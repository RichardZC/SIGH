VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEconRepSIS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes SIS"
   ClientHeight    =   10440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13530
   Icon            =   "EconRepSis.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10440
   ScaleWidth      =   13530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab 
      Height          =   10335
      Left            =   15
      TabIndex        =   0
      Top             =   45
      Width           =   13440
      _ExtentX        =   23707
      _ExtentY        =   18230
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
      TabCaption(0)   =   "Consumo individual por Servicios"
      TabPicture(0)   =   "EconRepSis.frx":0CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(4)=   "Label18"
      Tab(0).Control(5)=   "txtFechaFin"
      Tab(0).Control(6)=   "txtFechaInicio"
      Tab(0).Control(7)=   "optFechas"
      Tab(0).Control(8)=   "optTodos"
      Tab(0).Control(9)=   "grdAtenciones"
      Tab(0).Control(10)=   "grdServicios"
      Tab(0).Control(11)=   "cmdBuscaCuentaPorApellidos"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame1"
      Tab(0).Control(13)=   "txtCA"
      Tab(0).Control(14)=   "txtAN"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtHC"
      Tab(0).Control(16)=   "Frame3"
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Impresión de FUAs ya grabadas de Consultorios Externos"
      TabPicture(1)   =   "EconRepSis.frx":0CE6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label9"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label10"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "grdFUAs"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtFecFin"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtFecInicio"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Frame2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmbIdServicioCE"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdExcel"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdEnviaEmail"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      Begin VB.CommandButton cmdEnviaEmail 
         Caption         =   "..."
         Height          =   450
         Left            =   12945
         TabIndex        =   46
         ToolTipText     =   "Envía EMAIL a cada Paciente porque tiene SIS VENCIDO"
         Top             =   1740
         Width           =   405
      End
      Begin VB.CommandButton cmdExcel 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   12945
         Picture         =   "EconRepSis.frx":0D02
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   1245
         Width           =   405
      End
      Begin VB.ComboBox cmbIdServicioCE 
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
         Left            =   1500
         TabIndex        =   41
         Text            =   "cmbIdServicioCE"
         Top             =   765
         Width           =   3705
      End
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   45
         TabIndex        =   35
         Top             =   9285
         Width           =   13350
         Begin VB.CommandButton btnImprimir 
            Caption         =   "Imprime c/FUA"
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
            Left            =   4950
            Picture         =   "EconRepSis.frx":11DB
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   165
            Width           =   1365
         End
         Begin VB.CommandButton cmdSalir 
            Cancel          =   -1  'True
            Caption         =   "Cancelar (ESC)"
            DisabledPicture =   "EconRepSis.frx":16B4
            DownPicture     =   "EconRepSis.frx":1B78
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
            Left            =   6390
            Picture         =   "EconRepSis.frx":2064
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   165
            Width           =   1365
         End
      End
      Begin VB.Frame Frame3 
         Height          =   975
         Left            =   -74970
         TabIndex        =   21
         Top             =   9300
         Width           =   13350
         Begin VB.CommandButton btnCancelar 
            Caption         =   "Cancelar (ESC)"
            DisabledPicture =   "EconRepSis.frx":2550
            DownPicture     =   "EconRepSis.frx":2A14
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
            Left            =   6390
            Picture         =   "EconRepSis.frx":2F00
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   165
            Width           =   1365
         End
         Begin VB.CommandButton btnAceptar 
            Caption         =   "Visualizar (F2)"
            DisabledPicture =   "EconRepSis.frx":33EC
            DownPicture     =   "EconRepSis.frx":384C
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
            Height          =   700
            Left            =   4860
            Picture         =   "EconRepSis.frx":3CC1
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   165
            Width           =   1365
         End
         Begin VB.CheckBox chkExcel 
            Alignment       =   1  'Right Justify
            Caption         =   "Exportar a Excel"
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
            Left            =   2640
            Picture         =   "EconRepSis.frx":4136
            TabIndex        =   22
            Top             =   360
            Visible         =   0   'False
            Width           =   1845
         End
      End
      Begin VB.TextBox txtHC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   -74085
         TabIndex        =   20
         Top             =   360
         Width           =   1260
      End
      Begin VB.TextBox txtAN 
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   -70455
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   360
         Width           =   5220
      End
      Begin VB.TextBox txtCA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   -63195
         TabIndex        =   18
         Top             =   360
         Width           =   1500
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   4695
         Left            =   -74970
         TabIndex        =   3
         Top             =   4620
         Width           =   13335
         Begin VB.TextBox txtNC 
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
            Height          =   315
            Left            =   90
            TabIndex        =   7
            Top             =   600
            Width           =   1740
         End
         Begin VB.TextBox txtD 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1275
            Left            =   90
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   1320
            Width           =   1740
         End
         Begin VB.TextBox txtFA 
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
            Height          =   315
            Left            =   90
            TabIndex        =   5
            Top             =   3000
            Width           =   1740
         End
         Begin VB.ComboBox cboServicio 
            Height          =   315
            Left            =   8400
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   0
            Visible         =   0   'False
            Width           =   4695
         End
         Begin TabDlg.SSTab sstReporte 
            Height          =   4215
            Left            =   1920
            TabIndex        =   8
            Top             =   360
            Width           =   11325
            _ExtentX        =   19976
            _ExtentY        =   7435
            _Version        =   393216
            Tabs            =   5
            TabsPerRow      =   5
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
            TabCaption(0)   =   "Medicamentos"
            TabPicture(0)   =   "EconRepSis.frx":4448
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "ssMed"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Insumos"
            TabPicture(1)   =   "EconRepSis.frx":4464
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "ssIns"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Laboratorio"
            TabPicture(2)   =   "EconRepSis.frx":4480
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "ssLab"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Rayos X"
            TabPicture(3)   =   "EconRepSis.frx":449C
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "ssImag"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "Procedimientos"
            TabPicture(4)   =   "EconRepSis.frx":44B8
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "ssProc"
            Tab(4).ControlCount=   1
            Begin UltraGrid.SSUltraGrid ssMed 
               Height          =   3405
               Left            =   120
               TabIndex        =   9
               Top             =   480
               Width           =   11055
               _ExtentX        =   19500
               _ExtentY        =   6006
               _Version        =   131072
               GridFlags       =   17040384
               LayoutFlags     =   67108884
               BorderStyle     =   5
               ScrollBars      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Medicamentos que consumió"
            End
            Begin UltraGrid.SSUltraGrid ssIns 
               Height          =   3405
               Left            =   -74880
               TabIndex        =   10
               Top             =   480
               Width           =   11055
               _ExtentX        =   19500
               _ExtentY        =   6006
               _Version        =   131072
               GridFlags       =   17040384
               LayoutFlags     =   67108884
               BorderStyle     =   5
               ScrollBars      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Insumos que consumió"
            End
            Begin UltraGrid.SSUltraGrid ssLab 
               Height          =   3405
               Left            =   -74880
               TabIndex        =   11
               Top             =   480
               Width           =   11055
               _ExtentX        =   19500
               _ExtentY        =   6006
               _Version        =   131072
               GridFlags       =   17040384
               LayoutFlags     =   67108884
               BorderStyle     =   5
               ScrollBars      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Pruebas de Laboratorio que se realizó"
            End
            Begin UltraGrid.SSUltraGrid ssImag 
               Height          =   3405
               Left            =   -74880
               TabIndex        =   12
               Top             =   480
               Width           =   11055
               _ExtentX        =   19500
               _ExtentY        =   6006
               _Version        =   131072
               GridFlags       =   17040384
               LayoutFlags     =   67108884
               BorderStyle     =   5
               ScrollBars      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Imágenes de Rayos X que se tomó"
            End
            Begin UltraGrid.SSUltraGrid ssProc 
               Height          =   3405
               Left            =   -74880
               TabIndex        =   13
               Top             =   480
               Width           =   11055
               _ExtentX        =   19500
               _ExtentY        =   6006
               _Version        =   131072
               GridFlags       =   17040384
               LayoutFlags     =   67108884
               BorderStyle     =   5
               ScrollBars      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Procedimientos que consumió"
            End
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Servicio ..."
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
            TabIndex        =   17
            Top             =   30
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Alta"
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
            Left            =   60
            TabIndex        =   16
            Top             =   2790
            Width           =   1125
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Cama"
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
            Left            =   60
            TabIndex        =   15
            Top             =   390
            Width           =   435
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Diagnóstico"
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
            Left            =   60
            TabIndex        =   14
            Top             =   1110
            Width           =   930
         End
      End
      Begin VB.CommandButton cmdBuscaCuentaPorApellidos 
         Caption         =   "..."
         Height          =   285
         Left            =   -72825
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Busca Cuenta por Apellidos y Nombres"
         Top             =   360
         Width           =   315
      End
      Begin UltraGrid.SSUltraGrid grdServicios 
         Height          =   1605
         Left            =   -74970
         TabIndex        =   2
         Top             =   2940
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   2831
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BorderStyle     =   5
         ScrollBars      =   2
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Servicios donde estuvo el paciente en la Atención seleccionada"
      End
      Begin UltraGrid.SSUltraGrid grdAtenciones 
         Height          =   1605
         Left            =   -74970
         TabIndex        =   25
         Top             =   1260
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   2831
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BorderStyle     =   5
         ScrollBars      =   2
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Atenciones que tuvo el paciente"
      End
      Begin Threed.SSOption optTodos 
         Height          =   240
         Left            =   -74190
         TabIndex        =   26
         Top             =   690
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   423
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
         Caption         =   "Todos los consumos"
         Value           =   -1
      End
      Begin Threed.SSOption optFechas 
         Height          =   240
         Left            =   -74190
         TabIndex        =   27
         Top             =   930
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   423
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
         Caption         =   "Consumos por Fechas"
      End
      Begin MSMask.MaskEdBox txtFechaInicio 
         Height          =   315
         Left            =   -70470
         TabIndex        =   28
         Top             =   900
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaFin 
         Height          =   315
         Left            =   -66030
         TabIndex        =   29
         Top             =   900
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFecInicio 
         Height          =   315
         Left            =   1500
         TabIndex        =   37
         Top             =   390
         Width           =   1365
         _ExtentX        =   2408
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
      Begin MSMask.MaskEdBox txtFecFin 
         Height          =   315
         Left            =   3825
         TabIndex        =   38
         Top             =   405
         Width           =   1365
         _ExtentX        =   2408
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
      Begin UltraGrid.SSUltraGrid grdFUAs 
         Height          =   7965
         Left            =   300
         TabIndex        =   44
         Top             =   1245
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   14049
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BorderStyle     =   5
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "FUAS vencidos (no se imprimen)"
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Consultorio CE"
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
         Left            =   285
         TabIndex        =   42
         Top             =   810
         Width           =   1170
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha Citas"
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
         Left            =   285
         TabIndex        =   40
         Top             =   435
         Width           =   915
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   3645
         TabIndex        =   39
         Top             =   450
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Código de Afiliación"
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
         Left            =   -64830
         TabIndex        =   34
         Top             =   390
         Width           =   1575
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "H. Clínica"
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
         Left            =   -74910
         TabIndex        =   33
         Top             =   390
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Apellidos y Nombres"
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
         Left            =   -72150
         TabIndex        =   32
         Top             =   390
         Width           =   1635
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
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
         Left            =   -67590
         TabIndex        =   31
         Top             =   960
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
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
         Left            =   -71790
         TabIndex        =   30
         Top             =   960
         Visible         =   0   'False
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmEconRepSIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reporte para SIS
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim oPaciente As New Pacientes '
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_paciente As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
Dim mo_ReglasHoteleria As New SIGHNegocios.ReglasHoteleria
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminServHosp As New ReglasServiciosHosp
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_SisConsumoWeb As New SIGHNegocios.SisConsumoWeb
Dim ml_idUsuario As Long
Dim ml_idPaciente As Long
Dim ml_idServicio As Long
Dim ml_idAtencion As Long
Dim ml_idCtaAtencion As Long
Dim ml_IdCama As Long
Dim ml_FechaAlta As Date
Dim ml_Cama As String
Dim ml_Diagnostico As String
Dim ml_Historia As Long
Dim ml_Nombres As String
Dim ml_CAfiliacion As String
Dim ml_Servicio As String
Dim ml_Alta As String
Dim mo_cboServicio As New sighentidades.ListaDespleglable
Dim mo_cmbIdServicioCE As New sighentidades.ListaDespleglable
Dim gridInfra As New GridInfragistic

Dim oDOPaciente As New DOPaciente
Dim oDOCama As New DOCama
Dim oServicios As ADODB.Recordset
Dim oAtenciones As ADODB.Recordset
Dim oDiagnosticos As ADODB.Recordset
Dim oLab As ADODB.Recordset
Dim oImag As ADODB.Recordset
Dim oMed As ADODB.Recordset
Dim oIns As ADODB.Recordset
Dim oProc As ADODB.Recordset
Dim oDevolucion As ADODB.Recordset

Dim rsTmp As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes
Dim lcFuenteFinanciamientoElegida As String
  
Property Let idUsuario(lValue As Long)
  ml_idUsuario = lValue
End Property

Sub CargaDB_TextBox(Tabla As ADODB.Recordset, T As TextBox)
  Dim K As Integer
  T.Text = ""
  K = 0
  If Tabla.EOF = True And Tabla.BOF = True Then Exit Sub
  Tabla.MoveFirst
  Do While Not (Tabla.EOF)
    K = K + 1
    If K = 1 Then
      T.Text = K & " --> " & Tabla!CodigoCIE10
    Else
      T.Text = T.Text & vbCrLf & K & " --> " & Tabla!CodigoCIE10
    End If
    Tabla.MoveNext
  Loop
  Tabla.Close
End Sub

Private Function BuscaPaciente(HCPaciente As Long)
  If HCPaciente = 0 Then Exit Function
  Set oDOPaciente = mo_paciente.PacientesSeleccionarPorHistoriaClinicaDefinitiva(HCPaciente)
  ml_idPaciente = Val(oDOPaciente.idPaciente)
  If ml_idPaciente <> 0 Then
    ml_Nombres = oDOPaciente.ApellidoPaterno & " " & oDOPaciente.ApellidoMaterno & ", " & oDOPaciente.PrimerNombre & " " & oDOPaciente.SegundoNombre
    Set oAtenciones = mo_ReglasLaboratorio.AtencionesQueTuvoElPaciente(ml_idPaciente)
    Set grdAtenciones.DataSource = oAtenciones
    grdAtenciones.Enabled = True
  Else
    ml_Nombres = ""
    ml_idPaciente = 0
    grdAtenciones.Enabled = False
  End If
  txtAN.Text = ml_Nombres
End Function

Private Sub btnAceptar_Click()
  Dim iFila As Long, iCol As Integer
  Dim rsReporte As New Recordset
  Dim II As Integer, Devueltos As Long
  Dim TCant As Long, TPrec As Double, TotGen As Double
  Dim TCant1 As Long, TotGen1 As Double, Cod As String, TPrec1 As Double
  Dim lcBuscaParametro As New SIGHDatos.Parametros
  Dim mo_ReporteUtil As New ReporteUtil
  Dim lbEsOpenOffice As Boolean
  Dim lcSql As String
  
  lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)

    If lbEsOpenOffice = True Then
        Dim ServiceManager As Object
        Dim Desktop As Object
        Dim Document As Object
        Dim Feuille As Object
        Dim Plage As Object
        Dim args()
        Dim Chemin As String
        Dim Fichier As String
        Dim lcArchivoExcel As String
        Dim PrintArea(0)
        Dim Style As Object
        Dim Border As Object
        'encabezado
        Dim PageStyles As Object
        Dim Sheet As Object
        Dim StyleFamilies As Object
        Dim DefPage As Object
        Dim Htext As Object
        Dim Hcontent As Object
        Dim ret As Long
        Dim lnHwnd As Long
    Else
        Dim oExcel As Excel.Application
        Dim oWorkBookPlantilla As Workbook
        Dim oWorkBook As Workbook
        Dim oWorkSheet As Worksheet
    End If
  
  If optFechas.Value = True Then
    If Trim(txtFechaInicio.Text) = "" Or Not IsDate(txtFechaInicio.Text) Then
      MsgBox "Por favor ingrese la fecha inicial", vbInformation, Me.Caption
      txtFechaInicio.SetFocus
      Exit Sub
    End If
    If Trim(txtFechaFin.Text) = "" Or Not IsDate(txtFechaFin.Text) Then
      MsgBox "Por favor ingrese la fecha final", vbInformation, Me.Caption
      txtFechaFin.SetFocus
      Exit Sub
    End If
  End If
  
  MousePointer = 11
  TotGen = 0
  'Crea nueva hoja
  
      If lbEsOpenOffice = True Then
        'Abre el archivo ExcelOpenOffice
        lcArchivoExcel = App.Path + "\Plantillas\CoberturaSIS.ods"
'        FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
'        Chemin = "file:///" & App.Path & "\Plantillas\"
'        Chemin = Replace(Chemin, "\", "/")
'        Fichier = Chemin & "/OpenOffice.ods"
        Fichier = Format(Time, "hhmmss") & ".ods"
        FileCopy lcArchivoExcel, App.Path + "\Plantillas\" & Fichier
        lcArchivoExcel = Fichier
        Chemin = "file:///" & App.Path & "\Plantillas\"
        Chemin = Replace(Chemin, "\", "/")
        Fichier = Chemin & "/" & lcArchivoExcel
        '
        Set ServiceManager = CreateObject("com.sun.star.ServiceManager")
        Set Desktop = ServiceManager.createInstance("com.sun.star.frame.Desktop")
        Set Document = Desktop.loadComponentFromURL(Fichier, "_blank", 0, args)
        Set Feuille = Document.getSheets().getByIndex(0)
        mo_CabeceraReportes.CabeceraReportes Document, True
        ret = SetForegroundWindow(lnHwnd)
    Else
        Set oExcel = GalenhosExcelApplication()
        Set oWorkBook = oExcel.Workbooks.Add
        'Abre, copia y cierra la plantilla
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\CoberturaSIS.xls")
        oWorkBookPlantilla.Worksheets("SIS").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
  
        '------- MEDICINAS
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
    End If
  'Inicio de Impresion
    If lbEsOpenOffice = True Then
        'Call Feuille.getcellbyposition(6, 0).setFormula(lcBuscaParametro.RetornaFechaHoraServidorSQL)
        Call Feuille.getcellbyposition(3, 4).setFormula(Left(ml_Nombres, 65) & "   (" & Left(lcFuenteFinanciamientoElegida, 55) & ")")
        Call Feuille.getcellbyposition(3, 6).setFormula(ml_CAfiliacion)
        Call Feuille.getcellbyposition(3, 8).setFormula(ml_Servicio)
        Call Feuille.getcellbyposition(3, 10).setFormula(ml_Diagnostico)
        Call Feuille.getcellbyposition(6, 6).setFormula(ml_Cama)
        Call Feuille.getcellbyposition(6, 8).setFormula(ml_Historia)
        Call Feuille.getcellbyposition(6, 10).setFormula(ml_Alta)
    Else
       ' oWorkSheet.Cells(1, 7).Value = lcBuscaParametro.RetornaFechaHoraServidorSQL
        oWorkSheet.Cells(5, 4).Value = Left(ml_Nombres, 65) & "   (" & Left(lcFuenteFinanciamientoElegida, 55) & ")"
        oWorkSheet.Cells(7, 4).Value = ml_CAfiliacion
        oWorkSheet.Cells(9, 4).Value = ml_Servicio
        oWorkSheet.Cells(11, 4).Value = ml_Diagnostico
        oWorkSheet.Cells(7, 7).Value = ml_Cama
        oWorkSheet.Cells(9, 7).Value = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(str(ml_Historia)), True)
        oWorkSheet.Cells(11, 7).Value = ml_Alta
    End If
    iFila = 13
    iCol = 2
  
  If oMed.State = adStateOpen And Not (oMed.EOF = True And oMed.BOF = True) Then
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(1, iFila - 1).setFormula("MEDICINAS")
    Else
        oWorkSheet.Cells(iFila, 2).Value = "MEDICINAS"
    End If
  iFila = iFila + 1
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula("Nº")
        Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula("Código")
        Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula("Descripción")
        Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula("Cantidad")
        Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula("Prec. Unit.")
        Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula("Monto")
        Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula("Dx")
    Else
        oWorkSheet.Cells(iFila, iCol).Value = "Nº"
        oWorkSheet.Cells(iFila, iCol + 1).Value = "Código"
        oWorkSheet.Cells(iFila, iCol + 2).Value = "Descripción"
        oWorkSheet.Cells(iFila, iCol + 3).Value = "Cantidad"
        oWorkSheet.Cells(iFila, iCol + 4).Value = "Prec. Unit."
        oWorkSheet.Cells(iFila, iCol + 5).Value = "Monto"
        oWorkSheet.Cells(iFila, iCol + 6).Value = "Dx"
    End If
  iFila = iFila + 1
  oMed.MoveFirst
  II = 0: TCant = 0: TPrec = 0
  TCant1 = 0: TotGen1 = 0: Cod = "": TPrec1 = 0
  Do While Not oMed.EOF
    TCant1 = 0
    II = II + 1
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(II)
        Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(oMed!codigo)
    Else
        oWorkSheet.Cells(iFila, iCol).Value = II
        oWorkSheet.Cells(iFila, iCol + 1).Value = oMed!codigo
    End If
    Cod = oMed!codigo
    Devueltos = 0 ' AveriguaDevueltos1(Cod) 'debb-05/04/2011
    TPrec1 = oMed!precio
    Do While Cod = oMed!codigo And oMed.BOF = False And oMed.EOF = False
      'If oMed!precio = TPrec1 Then
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(oMed!nombre)
        Else
            oWorkSheet.Cells(iFila, iCol + 2).Value = oMed!nombre
        End If
        TCant = TCant + oMed!Cantidad - Devueltos
        TCant1 = TCant1 + oMed!Cantidad - Devueltos
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(TCant1)
            Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(oMed!precio)
        Else
            oWorkSheet.Cells(iFila, iCol + 3).Value = TCant1
            oWorkSheet.Cells(iFila, iCol + 4).Value = oMed!precio
        End If
        TotGen1 = TCant1 * oMed!precio
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(TotGen1)
            Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(1)
        Else
            oWorkSheet.Cells(iFila, iCol + 5).Value = TotGen1
            oWorkSheet.Cells(iFila, iCol + 6).Value = 1
        End If
        oMed.MoveNext
      'End If
      If oMed.EOF = True Then Exit Do
    Loop
    TPrec = TPrec + TotGen1
    iFila = iFila + 1
    If oMed.EOF = True Then Exit Do
  Loop
    If lbEsOpenOffice = True Then
    Else
        oWorkSheet.Range(oWorkSheet.Cells(iFila - II, 2), oWorkSheet.Cells(iFila - 1, 8)).Borders.LineStyle = 1
    End If
    If lbEsOpenOffice = True Then
        Set Plage = Feuille.getCellRangeByName("D" & CStr(iFila) & ":G" & CStr(iFila))
        mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
        Call Feuille.getcellbyposition(3, iFila - 1).setFormula("SUB TOTAL DE MEDICINAS")
        Call Feuille.getcellbyposition(4, iFila - 1).setFormula(TCant)
        Call Feuille.getcellbyposition(6, iFila - 1).setFormula(Format(TPrec, "0.00"))
    Else
        oWorkSheet.Cells(iFila, 4).Value = "SUB TOTAL DE MEDICINAS"
        oWorkSheet.Cells(iFila, 4).Borders.LineStyle = 1
        oWorkSheet.Cells(iFila, 4).Font.Bold = True
        oWorkSheet.Cells(iFila, 5).Value = TCant
        oWorkSheet.Cells(iFila, 5).Borders.LineStyle = 1
        oWorkSheet.Cells(iFila, 5).Font.Bold = True
        oWorkSheet.Cells(iFila, 7).Value = Format(TPrec, "0.00")
        oWorkSheet.Cells(iFila, 7).Borders.LineStyle = 1
        oWorkSheet.Cells(iFila, 7).Font.Bold = True
    End If
  iFila = iFila + 2
  TotGen = TotGen + TPrec
  End If
  
  '-INSUMOS
  iCol = 2
  If oIns.State = adStateOpen And Not (oIns.EOF = True And oIns.BOF = True) Then
    If lbEsOpenOffice = True Then
      Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":H" & CStr(iFila))
      Plage.Merge (True)
      Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":H" & CStr(iFila))
      mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
      Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila + 1) & ":H" & CStr(iFila + 1))
      mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
      Call Feuille.getcellbyposition(1, iFila - 1).setFormula("INSUMOS")
    Else
      oWorkSheet.Cells(iFila, 2).Value = "INSUMOS"
      oWorkSheet.Range(oWorkSheet.Cells(iFila, 2), oWorkSheet.Cells(iFila, 8)).Merge
      oWorkSheet.Range(oWorkSheet.Cells(iFila, 2), oWorkSheet.Cells(iFila, 8)).Font.Bold = True
      oWorkSheet.Cells(iFila, 2).HorizontalAlignment = -4108 'xlCenter
    End If
  iFila = iFila + 1
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula("Nº")
        Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula("Código")
        Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula("Descripción")
        Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula("Cantidad")
        Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula("Prec. Unit.")
        Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula("Monto")
        Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula("Dx")
    Else
        oWorkSheet.Cells(iFila, iCol).Value = "Nº"
        oWorkSheet.Cells(iFila, iCol + 1).Value = "Código"
        oWorkSheet.Cells(iFila, iCol + 2).Value = "Descripción"
        oWorkSheet.Cells(iFila, iCol + 3).Value = "Cantidad"
        oWorkSheet.Cells(iFila, iCol + 4).Value = "Prec. Unit."
        oWorkSheet.Cells(iFila, iCol + 5).Value = "Monto"
        oWorkSheet.Cells(iFila, iCol + 6).Value = "Dx"
        oWorkSheet.Range(oWorkSheet.Cells(iFila - 1, 2), oWorkSheet.Cells(iFila, 8)).Borders.LineStyle = 1
        oWorkSheet.Range(oWorkSheet.Cells(iFila - 1, 1), oWorkSheet.Cells(iFila, 8)).Font.Bold = True
    End If
  iFila = iFila + 1
  oIns.MoveFirst
  II = 0: TCant = 0: TPrec = 0
  TCant1 = 0: TotGen1 = 0: Cod = "": TPrec1 = 0
  Do While Not oIns.EOF
    TCant1 = 0
    II = II + 1
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(II)
        Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(oIns!codigo)
    Else
        oWorkSheet.Cells(iFila, iCol).Value = II
        oWorkSheet.Cells(iFila, iCol + 1).Value = oIns!codigo
    End If
    Cod = oIns!codigo
    Devueltos = 0   'AveriguaDevueltos1(Cod)   'debb-05/04/2011
    TPrec1 = oIns!precio
    Do While Cod = oIns!codigo And oIns.BOF = False And oIns.EOF = False
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(oIns!nombre)
        Else
            oWorkSheet.Cells(iFila, iCol + 2).Value = oIns!nombre
        End If
        TCant = TCant + oIns!Cantidad - Devueltos
        TCant1 = TCant1 + oIns!Cantidad - Devueltos
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(TCant1)
            Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(oIns!precio)
        Else
            oWorkSheet.Cells(iFila, iCol + 3).Value = TCant1
            oWorkSheet.Cells(iFila, iCol + 4).Value = oIns!precio
        End If
        TotGen1 = TCant1 * oIns!precio
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(TotGen1)
            Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(1)
        Else
            TotGen1 = TCant1 * oIns!precio
            oWorkSheet.Cells(iFila, iCol + 5).Value = TotGen1
            oWorkSheet.Cells(iFila, iCol + 6).Value = 1
        End If
        oIns.MoveNext
      If oIns.EOF = True Then Exit Do
    Loop
    TPrec = TPrec + TotGen1
    iFila = iFila + 1
    If oIns.EOF = True Then Exit Do
  Loop
    If lbEsOpenOffice = True Then
    Else
        oWorkSheet.Range(oWorkSheet.Cells(iFila - II, 2), oWorkSheet.Cells(iFila - 1, 8)).Borders.LineStyle = 1
    End If
    If lbEsOpenOffice = True Then
        Set Plage = Feuille.getCellRangeByName("D" & CStr(iFila) & ":G" & CStr(iFila))
        mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
        Call Feuille.getcellbyposition(3, iFila - 1).setFormula("SUB TOTAL DE INSUMOS")
        Call Feuille.getcellbyposition(4, iFila - 1).setFormula(TCant)
        Call Feuille.getcellbyposition(6, iFila - 1).setFormula(Format(TPrec, "0.00"))
    Else
        oWorkSheet.Cells(iFila, 4).Value = "SUB TOTAL DE INSUMOS"
        oWorkSheet.Cells(iFila, 4).Borders.LineStyle = 1
        oWorkSheet.Cells(iFila, 4).Font.Bold = True
        oWorkSheet.Cells(iFila, 5).Value = TCant
        oWorkSheet.Cells(iFila, 5).Borders.LineStyle = 1
        oWorkSheet.Cells(iFila, 5).Font.Bold = True
        oWorkSheet.Cells(iFila, 7).Value = Format(TPrec, "0.00")
        oWorkSheet.Cells(iFila, 7).Borders.LineStyle = 1
        oWorkSheet.Cells(iFila, 7).Font.Bold = True
    End If
  iFila = iFila + 2
  TotGen = TotGen + TPrec
  End If
  
  '-Procedimientos
  iCol = 2
  If oProc.State = adStateOpen And Not (oProc.EOF = True And oProc.BOF = True) Then
    If lbEsOpenOffice = True Then
        Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":H" & CStr(iFila))
        Plage.Merge (True)
        Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":H" & CStr(iFila))
        mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
        Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila + 1) & ":H" & CStr(iFila + 1))
        mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
        Call Feuille.getcellbyposition(1, iFila - 1).setFormula("PROCEDIMIENTOS")
    Else
        oWorkSheet.Cells(iFila, 2).Value = "PROCEDIMIENTOS"
        oWorkSheet.Range(oWorkSheet.Cells(iFila, 2), oWorkSheet.Cells(iFila, 8)).Merge
        oWorkSheet.Range(oWorkSheet.Cells(iFila, 2), oWorkSheet.Cells(iFila, 8)).Font.Bold = True
        oWorkSheet.Cells(iFila, 2).HorizontalAlignment = -4108 'xlCenter
    End If
  iFila = iFila + 1
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula("Nº")
        Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula("Código")
        Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula("Descripción")
        Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula("Cantidad")
        Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula("Prec. Unit.")
        Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula("Monto")
        Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula("Dx")
    Else
        oWorkSheet.Cells(iFila, iCol).Value = "Nº"
        oWorkSheet.Cells(iFila, iCol + 1).Value = "Código"
        oWorkSheet.Cells(iFila, iCol + 2).Value = "Descripción"
        oWorkSheet.Cells(iFila, iCol + 3).Value = "Cantidad"
        oWorkSheet.Cells(iFila, iCol + 4).Value = "Prec. Unit."
        oWorkSheet.Cells(iFila, iCol + 5).Value = "Monto"
        oWorkSheet.Cells(iFila, iCol + 6).Value = "Dx"
        oWorkSheet.Range(oWorkSheet.Cells(iFila - 1, 2), oWorkSheet.Cells(iFila, 8)).Borders.LineStyle = 1
        oWorkSheet.Range(oWorkSheet.Cells(iFila - 1, 1), oWorkSheet.Cells(iFila, 8)).Font.Bold = True
    End If
  iFila = iFila + 1
  oProc.MoveFirst
  II = 0: TCant = 0: TPrec = 0
  TCant1 = 0: TotGen1 = 0: Cod = "": TPrec1 = 0
  Do While Not oProc.EOF
    TCant1 = 0
    II = II + 1
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(II)
        Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(oProc!codigo)
    Else
        oWorkSheet.Cells(iFila, iCol).Value = II
        oWorkSheet.Cells(iFila, iCol + 1).Value = oProc!codigo
    End If
    Cod = oProc!codigo
    TPrec1 = oProc!precio
    Do While Cod = oProc!codigo And oProc.BOF = False And oProc.EOF = False
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(oProc!nombre)
        Else
            oWorkSheet.Cells(iFila, iCol + 2).Value = oProc!nombre
        End If
        TCant = TCant + oProc!Cantidad
        TCant1 = TCant1 + oProc!Cantidad
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(TCant1)
        Else
            oWorkSheet.Cells(iFila, iCol + 3).Value = TCant1
        End If
        If Trim(Cod) = "F00001" Then
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(oProc!precio * (-1))
            Else
                oWorkSheet.Cells(iFila, iCol + 4).Value = "'-" '-1 * oProc!Precio
            End If
          TotGen1 = TotGen1 + oProc!Cantidad * oProc!precio * (-1) ' TCant1 * oProc!Precio * (-1)
        Else
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(oProc!precio)
            Else
                oWorkSheet.Cells(iFila, iCol + 4).Value = oProc!precio
            End If
          TotGen1 = TCant1 * oProc!precio
        End If
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(TotGen1)
            Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(1)
        Else
            oWorkSheet.Cells(iFila, iCol + 5).Value = TotGen1
            oWorkSheet.Cells(iFila, iCol + 6).Value = 1
        End If
        oProc.MoveNext
      If oProc.EOF = True Then Exit Do
    Loop
    TPrec = TPrec + TotGen1
    iFila = iFila + 1
    If oProc.EOF = True Then Exit Do
  Loop
    If lbEsOpenOffice = True Then
    Else
        oWorkSheet.Range(oWorkSheet.Cells(iFila - II, 2), oWorkSheet.Cells(iFila - 1, 8)).Borders.LineStyle = 1
    End If
    If lbEsOpenOffice = True Then
        Set Plage = Feuille.getCellRangeByName("D" & CStr(iFila) & ":G" & CStr(iFila))
        mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
        Call Feuille.getcellbyposition(3, iFila - 1).setFormula("SUB TOTAL DE PROCEDIMIENTOS")
        Call Feuille.getcellbyposition(4, iFila - 1).setFormula(TCant)
        Call Feuille.getcellbyposition(6, iFila - 1).setFormula(Format(TPrec, "0.00"))
    Else
        oWorkSheet.Cells(iFila, 4).Value = "SUB TOTAL DE PROCEDIMIENTOS"
        oWorkSheet.Cells(iFila, 4).Borders.LineStyle = 1
        oWorkSheet.Cells(iFila, 4).Font.Bold = True
        oWorkSheet.Cells(iFila, 5).Value = TCant
        oWorkSheet.Cells(iFila, 5).Borders.LineStyle = 1
        oWorkSheet.Cells(iFila, 5).Font.Bold = True
        oWorkSheet.Cells(iFila, 7).Value = Format(TPrec, "0.00")
        oWorkSheet.Cells(iFila, 7).Borders.LineStyle = 1
        oWorkSheet.Cells(iFila, 7).Font.Bold = True
    End If
  iFila = iFila + 2
  TotGen = TotGen + TPrec
  End If
  
  '-Laboratorio
  iCol = 2
  If oLab.State = adStateOpen And Not (oLab.EOF = True And oLab.BOF = True) Then
    If lbEsOpenOffice = True Then
        Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":H" & CStr(iFila))
        Plage.Merge (True)
        Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":H" & CStr(iFila))
        mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
        Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila + 1) & ":H" & CStr(iFila + 1))
        mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
        Call Feuille.getcellbyposition(1, iFila - 1).setFormula("LABORATORIO")
    Else
        oWorkSheet.Cells(iFila, 2).Value = "LABORATORIO"
        oWorkSheet.Range(oWorkSheet.Cells(iFila, 2), oWorkSheet.Cells(iFila, 8)).Merge
        oWorkSheet.Range(oWorkSheet.Cells(iFila, 2), oWorkSheet.Cells(iFila, 8)).Font.Bold = True
        oWorkSheet.Cells(iFila, 2).HorizontalAlignment = -4108 'xlCenter
    End If
  iFila = iFila + 1
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula("Nº")
        Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula("Código")
        Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula("Descripción")
        Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula("Cantidad")
        Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula("Prec. Unit.")
        Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula("Monto")
        Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula("Dx")
    Else
        oWorkSheet.Cells(iFila, iCol).Value = "Nº"
        oWorkSheet.Cells(iFila, iCol + 1).Value = "Código"
        oWorkSheet.Cells(iFila, iCol + 2).Value = "Descripción"
        oWorkSheet.Cells(iFila, iCol + 3).Value = "Cantidad"
        oWorkSheet.Cells(iFila, iCol + 4).Value = "Prec. Unit."
        oWorkSheet.Cells(iFila, iCol + 5).Value = "Monto"
        oWorkSheet.Cells(iFila, iCol + 6).Value = "Dx"
        oWorkSheet.Range(oWorkSheet.Cells(iFila - 1, 2), oWorkSheet.Cells(iFila, 8)).Borders.LineStyle = 1
        oWorkSheet.Range(oWorkSheet.Cells(iFila - 1, 1), oWorkSheet.Cells(iFila, 8)).Font.Bold = True
    End If
  iFila = iFila + 1
  oLab.MoveFirst
  II = 0: TCant = 0: TPrec = 0
  TCant1 = 0: TotGen1 = 0: Cod = "": TPrec1 = 0
  Do While Not oLab.EOF
    TCant1 = 0
    II = II + 1
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(II)
        Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(oLab!codigo)
    Else
        oWorkSheet.Cells(iFila, iCol).Value = II
        oWorkSheet.Cells(iFila, iCol + 1).Value = oLab!codigo
    End If
    Cod = oLab!codigo
    TPrec1 = oLab!precio
    Do While Cod = oLab!codigo And oLab.BOF = False And oLab.EOF = False
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(oLab!nombre)
        Else
            oWorkSheet.Cells(iFila, iCol + 2).Value = oLab!nombre
        End If
        TCant = TCant + oLab!Cantidad
        TCant1 = TCant1 + oLab!Cantidad
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(TCant1)
            Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(oLab!precio)
        Else
            oWorkSheet.Cells(iFila, iCol + 3).Value = TCant1
            oWorkSheet.Cells(iFila, iCol + 4).Value = oLab!precio
        End If
        TotGen1 = TCant1 * oLab!precio
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(TotGen1)
            Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(1)
        Else
            oWorkSheet.Cells(iFila, iCol + 5).Value = TotGen1
            oWorkSheet.Cells(iFila, iCol + 6).Value = 1
        End If
      oLab.MoveNext
      If oLab.EOF = True Then Exit Do
    Loop
    TPrec = TPrec + TotGen1
    iFila = iFila + 1
    If oLab.EOF = True Then Exit Do
  Loop
    If lbEsOpenOffice = True Then
    Else
        oWorkSheet.Range(oWorkSheet.Cells(iFila - II, 2), oWorkSheet.Cells(iFila - 1, 8)).Borders.LineStyle = 1
    End If
    If lbEsOpenOffice = True Then
        Set Plage = Feuille.getCellRangeByName("D" & CStr(iFila) & ":G" & CStr(iFila))
        mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
        Call Feuille.getcellbyposition(3, iFila - 1).setFormula("SUB TOTAL DE LABORATORIO")
        Call Feuille.getcellbyposition(4, iFila - 1).setFormula(TCant)
        Call Feuille.getcellbyposition(6, iFila - 1).setFormula(Format(TPrec, "0.00"))
    Else
        oWorkSheet.Cells(iFila, 4).Value = "SUB TOTAL DE LABORATORIO"
        oWorkSheet.Cells(iFila, 4).Borders.LineStyle = 1
        oWorkSheet.Cells(iFila, 4).Font.Bold = True
        oWorkSheet.Cells(iFila, 5).Value = TCant
        oWorkSheet.Cells(iFila, 5).Borders.LineStyle = 1
        oWorkSheet.Cells(iFila, 5).Font.Bold = True
        oWorkSheet.Cells(iFila, 7).Value = Format(TPrec, "0.00")
        oWorkSheet.Cells(iFila, 7).Borders.LineStyle = 1
        oWorkSheet.Cells(iFila, 7).Font.Bold = True
    End If
  iFila = iFila + 2
  TotGen = TotGen + TPrec
  End If
  
  '-Rayos X
  iCol = 2
  If oImag.State = adStateOpen And Not (oImag.EOF = True And oImag.BOF = True) Then
    If lbEsOpenOffice = True Then
        Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":H" & CStr(iFila))
        Plage.Merge (True)
        Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":H" & CStr(iFila))
        mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
        Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila + 1) & ":H" & CStr(iFila + 1))
        mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
        Call Feuille.getcellbyposition(1, iFila - 1).setFormula("IMÁGENES")
    Else
        oWorkSheet.Cells(iFila, 2).Value = "IMÁGENES"
        oWorkSheet.Range(oWorkSheet.Cells(iFila, 2), oWorkSheet.Cells(iFila, 8)).Merge
        oWorkSheet.Range(oWorkSheet.Cells(iFila, 2), oWorkSheet.Cells(iFila, 8)).Font.Bold = True
        oWorkSheet.Cells(iFila, 2).HorizontalAlignment = -4108 'xlCenter
    End If
  iFila = iFila + 1
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula("Nº")
        Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula("Código")
        Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula("Descripción")
        Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula("Cantidad")
        Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula("Prec. Unit.")
        Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula("Monto")
        Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula("Dx")
    Else
        oWorkSheet.Cells(iFila, iCol).Value = "Nº"
        oWorkSheet.Cells(iFila, iCol + 1).Value = "Código"
        oWorkSheet.Cells(iFila, iCol + 2).Value = "Descripción"
        oWorkSheet.Cells(iFila, iCol + 3).Value = "Cantidad"
        oWorkSheet.Cells(iFila, iCol + 4).Value = "Prec. Unit."
        oWorkSheet.Cells(iFila, iCol + 5).Value = "Monto"
        oWorkSheet.Cells(iFila, iCol + 6).Value = "Dx"
        oWorkSheet.Range(oWorkSheet.Cells(iFila - 1, 2), oWorkSheet.Cells(iFila, 8)).Borders.LineStyle = 1
        oWorkSheet.Range(oWorkSheet.Cells(iFila - 1, 1), oWorkSheet.Cells(iFila, 8)).Font.Bold = True
    End If
  iFila = iFila + 1
  oImag.MoveFirst
  II = 0: TCant = 0: TPrec = 0
  TCant1 = 0: TotGen1 = 0: Cod = "": TPrec1 = 0
  Do While Not oImag.EOF
    TCant1 = 0
    II = II + 1
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(II)
        Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(oImag!codigo)
    Else
        oWorkSheet.Cells(iFila, iCol).Value = II
        oWorkSheet.Cells(iFila, iCol + 1).Value = oImag!codigo
    End If
    Cod = oImag!codigo
    TPrec1 = oImag!precio
    Do While Cod = oImag!codigo And oImag.BOF = False And oImag.EOF = False
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(oImag!nombre)
        Else
            oWorkSheet.Cells(iFila, iCol + 2).Value = oImag!nombre
        End If
        TCant = TCant + oImag!Cantidad
        TCant1 = TCant1 + oImag!Cantidad
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(TCant1)
            Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(oImag!precio)
        Else
            oWorkSheet.Cells(iFila, iCol + 3).Value = TCant1
            oWorkSheet.Cells(iFila, iCol + 4).Value = oImag!precio
        End If
        TotGen1 = TCant1 * oImag!precio
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(TotGen1)
            Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(1)
        Else
            oWorkSheet.Cells(iFila, iCol + 5).Value = TotGen1
            oWorkSheet.Cells(iFila, iCol + 6).Value = 1
        End If
      oImag.MoveNext
      If oImag.EOF = True Then Exit Do
    Loop
Sigue:
    TPrec = TPrec + TotGen1
    iFila = iFila + 1
    If oImag.EOF = True Then Exit Do
  Loop
    If lbEsOpenOffice = True Then
    Else
          oWorkSheet.Range(oWorkSheet.Cells(iFila - II, 2), oWorkSheet.Cells(iFila - 1, 8)).Borders.LineStyle = 1
    End If
    If lbEsOpenOffice = True Then
        Set Plage = Feuille.getCellRangeByName("D" & CStr(iFila) & ":G" & CStr(iFila))
        mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
        Call Feuille.getcellbyposition(3, iFila - 1).setFormula("SUB TOTAL DE IMÁGENES")
        Call Feuille.getcellbyposition(4, iFila - 1).setFormula(TCant)
        Call Feuille.getcellbyposition(6, iFila - 1).setFormula(Format(TPrec, "0.00"))
    Else
        oWorkSheet.Cells(iFila, 4).Value = "SUB TOTAL DE IMÁGENES"
        oWorkSheet.Cells(iFila, 4).Borders.LineStyle = 1
        oWorkSheet.Cells(iFila, 4).Font.Bold = True
        oWorkSheet.Cells(iFila, 5).Value = TCant
        oWorkSheet.Cells(iFila, 5).Borders.LineStyle = 1
        oWorkSheet.Cells(iFila, 5).Font.Bold = True
        oWorkSheet.Cells(iFila, 7).Value = Format(TPrec, "0.00")
        oWorkSheet.Cells(iFila, 7).Borders.LineStyle = 1
        oWorkSheet.Cells(iFila, 7).Font.Bold = True
    End If
  iFila = iFila + 1
  TotGen = TotGen + TPrec
  End If
  
  iFila = iFila + 1
    If lbEsOpenOffice = True Then
        Set Plage = Feuille.getCellRangeByName("D" & CStr(iFila) & ":G" & CStr(iFila))
        mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
        Call Feuille.getcellbyposition(3, iFila - 1).setFormula("TOTAL GENERAL")
        Call Feuille.getcellbyposition(6, iFila - 1).setFormula(Format(TotGen, "0.00"))
    Else
        oWorkSheet.Cells(iFila, 4).Value = "TOTAL GENERAL"
        oWorkSheet.Cells(iFila, 4).Borders.LineStyle = 1
        oWorkSheet.Cells(iFila, 4).Font.Bold = True
        oWorkSheet.Cells(iFila, 7).Value = Format(TotGen, "0.00")
        oWorkSheet.Cells(iFila, 7).Borders.LineStyle = 1
        oWorkSheet.Cells(iFila, 7).Font.Bold = True
    End If
  iFila = iFila + 1
    If lbEsOpenOffice = True Then
        Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
        PrintArea(0).Sheet = 0
        PrintArea(0).startcolumn = 1
        PrintArea(0).StartRow = 0
        PrintArea(0).EndColumn = 8
        PrintArea(0).EndRow = iFila
        Call Feuille.SetPrintAreas(PrintArea())
        Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
        MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
    Else
        oWorkSheet.PageSetup.PrintTitleRows = "$1:$12"
        If oWorkSheet.PageSetup.PrintArea <> "" Then oWorkSheet.PageSetup.PrintArea = "$A$1:$I$" & (iFila + 2)
        oExcel.Visible = True
        oWorkSheet.PrintPreview
        MousePointer = 1
    End If
    If lbEsOpenOffice = True Then
        'Liberar Memoria
        Set Plage = Nothing
        Set Feuille = Nothing
        Set Document = Nothing
        Set Desktop = Nothing
        Set ServiceManager = Nothing
        Set Style = Nothing
        Set Border = Nothing
        'Liberar Memoria
        Set PageStyles = Nothing
        Set Sheet = Nothing
        Set StyleFamilies = Nothing
        Set DefPage = Nothing
        Set Htext = Nothing
        Set Hcontent = Nothing
    Else
        'Liberar memoria
        Set oExcel = Nothing
        Set oWorkBookPlantilla = Nothing
        Set oWorkBook = Nothing
        Set oWorkSheet = Nothing
    End If
  Exit Sub
  
  Dim oRptClaseCry As New CrystalR
  With rsTmp
    .Fields.Append "tipo", adVarChar, 50, adFldIsNullable
    .Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
    .Fields.Append "Nombre", adVarChar, 50, adFldIsNullable
    .Fields.Append "Cantidad", adInteger
    .Fields.Append "Precio", adDouble
    .Fields.Append "Total", adDouble
    .Fields.Append "idProducto", adInteger
    .LockType = adLockOptimistic
    .Open
  End With
  If Not (oMed.EOF = True And oMed.BOF = True) Then
    oMed.MoveFirst
    Do While Not oMed.EOF
      rsTmp.AddNew
      rsTmp!nombrep = Left(ml_Nombres, 100)
      rsTmp!Codigop = ml_CAfiliacion
      rsTmp!cama = ml_Cama
      rsTmp!Servicio = ml_Servicio
      rsTmp!HC = ml_Historia
      rsTmp!dx = ml_Diagnostico
      rsTmp!Alta = ml_Alta
      rsTmp!tipo = oMed!tipo
      rsTmp!codigo = oMed!codigo
      rsTmp!nombre = Left(oMed!nombre, 50)
      rsTmp!Cantidad = oMed!Cantidad
      rsTmp!precio = oMed!precio
      rsTmp!Total = oMed!Total
      rsTmp!idProducto = 1
      rsTmp.Update
      oMed.MoveNext
    Loop
  End If
  If Not (oIns.EOF = True And oIns.BOF = True) Then
    oIns.MoveFirst
    Do While Not oIns.EOF
      rsTmp.AddNew
      rsTmp!nombrep = Left(ml_Nombres, 100)
      rsTmp!Codigop = ml_CAfiliacion
      rsTmp!cama = ml_Cama
      rsTmp!Servicio = ml_Servicio
      rsTmp!HC = ml_Historia
      rsTmp!dx = ml_Diagnostico
      rsTmp!Alta = ml_Alta
      rsTmp!tipo = oIns!tipo
      rsTmp!codigo = oIns!codigo
      rsTmp!nombre = Left(oIns!nombre, 50)
      rsTmp!Cantidad = oIns!Cantidad
      rsTmp!precio = oIns!precio
      rsTmp!Total = oIns!Total
      rsTmp!idProducto = 1
      rsTmp.Update
      oIns.MoveNext
    Loop
  End If
  If Not (oLab.EOF = True And oLab.BOF = True) Then
    oLab.MoveFirst
    Do While Not oLab.EOF
      rsTmp.AddNew
      rsTmp!nombrep = Left(ml_Nombres, 100)
      rsTmp!Codigop = ml_CAfiliacion
      rsTmp!cama = ml_Cama
      rsTmp!Servicio = ml_Servicio
      rsTmp!HC = ml_Historia
      rsTmp!dx = ml_Diagnostico
      rsTmp!Alta = ml_Alta
      rsTmp!tipo = oLab!tipo
      rsTmp!codigo = oLab!codigo
      rsTmp!nombre = Left(oLab!nombre, 50)
      rsTmp!Cantidad = oLab!Cantidad
      rsTmp!precio = oLab!precio
      rsTmp!Total = oLab!Total
      rsTmp!idProducto = 1
      rsTmp.Update
      oLab.MoveNext
    Loop
  End If
  If Not (oImag.EOF = True And oImag.BOF = True) Then
    oImag.MoveFirst
    Do While Not oImag.EOF
      rsTmp.AddNew
      rsTmp!nombrep = Left(ml_Nombres, 100)
      rsTmp!Codigop = ml_CAfiliacion
      rsTmp!cama = ml_Cama
      rsTmp!Servicio = ml_Servicio
      rsTmp!HC = ml_Historia
      rsTmp!dx = ml_Diagnostico
      rsTmp!Alta = ml_Alta
      rsTmp!tipo = oImag!tipo
      rsTmp!codigo = oImag!codigo
      rsTmp!nombre = Left(oImag!nombre, 50)
      rsTmp!Cantidad = oImag!Cantidad
      rsTmp!precio = oImag!precio
      rsTmp!Total = oImag!Total
      rsTmp!idProducto = 1
      rsTmp.Update
      oImag.MoveNext
    Loop
  End If
  If Not (oProc.EOF = True And oProc.BOF = True) Then
    oProc.MoveFirst
    Do While Not oProc.EOF
      rsTmp.AddNew
      rsTmp!nombrep = Left(ml_Nombres, 100)
      rsTmp!Codigop = ml_CAfiliacion
      rsTmp!cama = ml_Cama
      rsTmp!Servicio = ml_Servicio
      rsTmp!HC = ml_Historia
      rsTmp!dx = ml_Diagnostico
      rsTmp!Alta = ml_Alta
      rsTmp!tipo = oProc!tipo
      rsTmp!codigo = oProc!codigo
      rsTmp!nombre = Left(oProc!nombre, 50)
      rsTmp!Cantidad = oProc!Cantidad
      rsTmp!precio = oProc!precio
      rsTmp!Total = oProc!Total
      rsTmp!idProducto = 1
      rsTmp.Update
      oProc.MoveNext
    Loop
  End If
  If rsTmp.EOF = True And rsTmp.BOF = True Then
    MsgBox "No se encuentran datos para mostrar", vbInformation, "SIGH"
  Else
    oRptClaseCry.Excel = IIf(chkExcel.Value = 1, True, False)
    oRptClaseCry.Archivo = "EconRepSIS"
    oRptClaseCry.Tabla = rsTmp
    oRptClaseCry.Show vbModal
    Set oRptClaseCry = Nothing
    Set rsTmp = Nothing
  End If

End Sub

Private Sub btnCancelar_Click()
  Unload Me
End Sub

Private Sub btnImprimir_Click()
    Dim lcMensajeLicencia As String
'    If False Then   'licencia
'       Exit Sub
'    End If
    
    Dim lcNombrePc As String, lnIdUsuario As Long
    Dim oFua As New SIGHSis.clFUA
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp76 As New Recordset
    Dim oRsTmp77 As New Recordset
    Dim oRsFuas As New Recordset
    Dim oRsBuscaPacientesSis As New Recordset
    Dim oConexionExterna As New Connection
    Dim mo_ReglasSISgalenhos As New SIGHSis.ReglasSISgalenhos
    Dim lcParametro322a As String, wxParametro323a As String, LbSisOK As Boolean
    Dim wxParametroJAMO1 As String, lcSql As String
    wxParametro323a = lcBuscaParametro.SeleccionaFilaParametro(323)
    lcParametro322a = lcBuscaParametro.SeleccionaFilaParametro(322)
    wxParametroJAMO1 = lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    oConexionExterna.CommandTimeout = 300
    oConexionExterna.CursorLocation = adUseClient
    oConexionExterna.Open wxParametroJAMO1
    
    With oRsFuas
          .Fields.Append "NroHistoria", adInteger, 4, adFldIsNullable
          .Fields.Append "Paciente", adVarChar, 100, adFldIsNullable
          .Fields.Append "Fua", adVarChar, 30, adFldIsNullable
          .Fields.Append "Cuenta", adInteger, 4, adFldIsNullable
          .Fields.Append "Cita", adVarChar, 20, adFldIsNullable
          .Fields.Append "Consultorio", adVarChar, 100, adFldIsNullable
          .Fields.Append "Email", adVarChar, 200, adFldIsNullable
          .Fields.Append "Telefono", adVarChar, 100, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
    Set Me.grdFUAs.DataSource = oRsFuas
    mo_Apariencia.ConfigurarFilasBiColores Me.grdFUAs, sighentidades.GrillaConFilasBicolor
    
    
    lcNombrePc = sighentidades.RetornaNombrePC
    lnIdUsuario = sighentidades.Usuario
    oFua.SoloImprimeFUAyaGrabado = True
    oFua.lcNombrePc = lcNombrePc
    oFua.lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroCitaCE
    oFua.idUsuario = lnIdUsuario
    oFua.Opcion = sghConsultar
    Set oRsTmp1 = mo_ReglasFacturacion.AtencionesSeleccionarCEPorFechaIngresoYplan(CDate(Me.txtFecInicio.Text), _
                                                                         CDate(Me.txtFecFin.Text), _
                                                                         sghFuenteFinanciamiento.sghFFSIS)
    If Val(mo_cmbIdServicioCE.BoundText) > 0 Then
       oRsTmp1.Filter = "idServicioIngreso=" & mo_cmbIdServicioCE.BoundText
    End If
    If oRsTmp1.RecordCount > 0 Then
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
          If Not IsNull(oRsTmp1!idSiaSis) And Not IsNull(oRsTmp1!SisCodigo) Then
            Set oRsTmp76 = mo_ReglasSISgalenhos.sisFuaAtencionXCuenta(oRsTmp1!idCuentaAtencion)
            If oRsTmp76.RecordCount > 0 Then
                '***********************Pacientes con ACREDITACION SIS VENCIDA (inicio)****************
                LbSisOK = False
                If lcParametro322a = "S" Then
                   '******* busca en la WEB
                   lcSql = ""
                   If Not IsNull(oRsTmp76!codigoTipoFormato) Then
                      lcSql = oRsTmp76!afiliacionNroIntegrante
                   End If
                   Set oRsBuscaPacientesSis = mo_SisConsumoWeb.WebServiceSISBuscarAfiliado("", _
                                                               oRsTmp76!AfiliacionDisa, oRsTmp76!AfiliacionTipoFormato, _
                                                               oRsTmp76!AfiliacionNroFormato, lcSql, _
                                                               oRsTmp76!codigo, wxParametro323a)
                   If oRsBuscaPacientesSis.RecordCount > 0 Then
                      If mo_ReglasSISgalenhos.Sis_ValidaSiEsAfiliadoActualDelSIS(oRsBuscaPacientesSis, _
                                                                   oRsTmp1!FechaIngreso, False) = True Then
                         LbSisOK = True
                          
                      End If
                   End If
                   oRsBuscaPacientesSis.Close
                End If
                If LbSisOK = False Then
                   '******** busca en la bd sigh_Externa
                   lcSql = "  where idSiaSis=" & oRsTmp76!idSiaSis & _
                            " and codigo='" & oRsTmp76!codigo & "'"
                   Set oRsBuscaPacientesSis = mo_ReglasSISgalenhos.SisFiltraPacientesAfiliados(lcSql, wxParametroJAMO1)
                   If oRsBuscaPacientesSis.RecordCount > 0 Then
                        If mo_ReglasSISgalenhos.Sis_ValidaSiEsAfiliadoActualDelSIS(oRsBuscaPacientesSis, _
                                                                      oRsTmp1!FechaIngreso, False) = True Then
                            LbSisOK = True
                        End If
                   End If
                   oRsBuscaPacientesSis.Close
                End If
                If LbSisOK = False Then
                   oRsFuas.AddNew
                   oRsFuas.Fields!NroHistoria = oRsTmp1!NroHistoriaClinica
                   oRsFuas.Fields!Paciente = oRsTmp1!Paciente
                   oRsFuas.Fields!fua = oRsTmp76!fuaDisa & "-" & oRsTmp76!fuaLote & "-" & oRsTmp76!fuaNumero
                   oRsFuas.Fields!cuenta = oRsTmp1!idCuentaAtencion
                   oRsFuas.Fields!Cita = Format(oRsTmp1!FechaIngreso, sighentidades.DevuelveFechaSoloFormato_DMY) & _
                                         " " & oRsTmp1!HoraIngreso
                   oRsFuas.Fields!Consultorio = oRsTmp1!DServicio
                   oRsFuas.Fields!email = IIf(IsNull(oRsTmp1!email), "", oRsTmp1!email)
                   oRsFuas.Fields!telefono = IIf(IsNull(oRsTmp1!telefono), "", oRsTmp1!telefono)
                   oRsFuas.Update
                   
                End If
                '***********************Pacientes con ACREDITACION SIS VENCIDA (fin)****************
                '
                If LbSisOK = True Then
                    oFua.idCuentaAtencion = oRsTmp1!idCuentaAtencion
                    oFua.idServicio = oRsTmp1!IdServicioIngreso
                    oFua.MostrarFormulario
                End If
            End If
            oRsTmp76.Close
           End If
           oRsTmp1.MoveNext
       Loop
    End If
    oRsTmp1.Close
    oConexionExterna.Close
    Set oFua = Nothing
    Set oRsTmp1 = Nothing
    Set oConexionExterna = Nothing
    Set oRsTmp76 = Nothing
    Set mo_ReglasSISgalenhos = Nothing
    Set oRsFuas = Nothing
    Set oRsBuscaPacientesSis = Nothing
End Sub

Private Sub cmdEnviaEmail_Click()
      Dim oMail As New SIGHNegocios.clsCDOmail
      Dim oRsTmp111 As New Recordset
      Dim lcEESS As String
      lcEESS = lcBuscaParametro.SeleccionaFilaParametro(205)
      Set oRsTmp111 = grdFUAs.DataSource
      If oRsTmp111.RecordCount > 0 Then
         oRsTmp111.MoveFirst
         Do While Not oRsTmp111.EOF
            If Not IsNull(oRsTmp111!email) Then
               oMail.EnviarUnSoloEmail "Problemas para la CITA del " & oRsTmp111!Cita & " en el " & lcEESS, _
                       "su SIS está vencido, debe acercarse al EESS para pagar por su CITA o CANCELARLA", _
                       "", oRsTmp111!email
            End If
            oRsTmp111.MoveNext
         Loop
      End If
      Set oMail = Nothing
      Set oRsTmp111 = Nothing
      MsgBox "Se envió Emails", vbInformation, ""
End Sub

Private Sub cmdExcel_Click()
    
    
    Dim oRsTmp1 As New Recordset
    Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
    Dim ml_TextoDelFiltro As String
    ml_TextoDelFiltro = Me.Label9.Caption & ": " & Me.txtFecInicio.Text & " al " & Me.txtFecFin.Text & _
                        IIf(Me.cmbIdServicioCE.Text = "", "", "   Consultorio:" & Me.cmbIdServicioCE.Text)
    Set oRsTmp1 = Me.grdFUAs.DataSource
    If oRsTmp1.State = 0 Then
       Exit Sub
    End If
    
    mo_AdminReportes.ExportarRecordSetAexcel oRsTmp1, Me.grdFUAs.Caption, ml_TextoDelFiltro, "", Me.hwnd, _
                                             True, True
    Set mo_AdminReportes = Nothing
    Set oRsTmp1 = Nothing
End Sub

Private Sub grdFUAs_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
  grdFUAs.Bands(0).Columns("NroHistoria").Width = 1000
  grdFUAs.Bands(0).Columns("Paciente").Width = 3000
  grdFUAs.Bands(0).Columns("Fua").Width = 1900
  grdFUAs.Bands(0).Columns("Cuenta").Width = 900
  grdFUAs.Bands(0).Columns("Cita").Width = 1800
  grdFUAs.Bands(0).Columns("Consultorio").Width = 3400
  grdFUAs.Bands(0).Columns("email").Hidden = True
  grdFUAs.Bands(0).Columns("telefono").Hidden = True
End Sub

Private Sub cmdBuscaCuentaPorApellidos_Click()
  Dim oBusqueda As New SIGHNegocios.BuscaPacientes
  Dim oDOPaciente As New DOPaciente
  Dim oConexion As New Connection
  oConexion.Open sighentidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  '
  oBusqueda.TipoFiltro = sghFiltrarTodos
  oBusqueda.MostrarFormulario
  If oBusqueda.BotonPresionado = sghAceptar Then
    Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(oBusqueda.IdRegistroSeleccionado, oConexion)
    If Not oDOPaciente Is Nothing Then
      txtHC.Text = oDOPaciente.NroHistoriaClinica
      txtHC.SetFocus
      SendKeys "{TAB}"
    End If
  End If
  oConexion.Close
  Set oConexion = Nothing
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Initialize()
  Set mo_cboServicio.MiComboBox = cboServicio
  Set mo_cmbIdServicioCE.MiComboBox = cmbIdServicioCE
End Sub

Private Sub Form_Load()
  ml_idServicio = 0
  ml_idPaciente = 0
  ml_idAtencion = 0
  mo_cboServicio.BoundColumn = "idServicio"
  mo_cboServicio.ListField = "NombreServicio"
  
  Me.txtFecInicio.Text = Format(Date + 1, sighentidades.DevuelveFechaSoloFormato_DMY)
  Me.txtFecFin.Text = Format(Date + 1, sighentidades.DevuelveFechaSoloFormato_DMY)
  mo_cmbIdServicioCE.BoundColumn = "idServicio"
  mo_cmbIdServicioCE.ListField = "descripcionLarga"
  Set mo_cmbIdServicioCE.RowSource = mo_AdminServHosp.ServiciosSeleccionarPorTipoV2(1, sghFiltraAnuladosYactivos)
  
End Sub

Private Sub grdAtenciones_Click()
  txtNC.Text = ""
  txtD.Text = ""
  txtFA.Text = ""
  Set ssLab.DataSource = Nothing
  Set ssImag.DataSource = Nothing
  Set ssIns.DataSource = Nothing
  Set ssMed.DataSource = Nothing
  Set ssProc.DataSource = Nothing
  
  'Busca servicios donde estuvo el paciente, por cada atención que tuvo
  If oAtenciones.EOF = True And oAtenciones.BOF = True Then Exit Sub
  'txtFechaInicio.Text = grdAtenciones.ActiveRow.Cells'if rsRecordset("IdLabEstado") = "0"
  ml_idAtencion = oAtenciones("idAtencion")
  ml_idCtaAtencion = oAtenciones("idCuentaAtencion")
  If oAtenciones!idTipoServicio = 1 Then
     Set oServicios = mo_ReglasLaboratorio.ServiciosDondeEstuvoElPacientece(ml_idAtencion)
  Else
     Set oServicios = mo_ReglasLaboratorio.ServiciosDondeEstuvoElPaciente(ml_idPaciente, ml_idAtencion)
  End If
  Set oDiagnosticos = mo_ReglasLaboratorio.DiagnosticosSeleccionarPorIdAtencion(ml_idAtencion)
  CargaDB_TextBox oDiagnosticos, txtD
  ml_Diagnostico = txtD.Text
  If Not (IsNull(oAtenciones("FechaEgreso"))) Then
    ml_Alta = oAtenciones("FechaEgreso")
  Else
    ml_Alta = ""
  End If
  txtFA.Text = ml_Alta
  txtFechaInicio.Text = oAtenciones("FechaIngreso") & " " & Format(oAtenciones("HoraIngreso"), "hh:mm:ss")
  txtFechaFin.Text = IIf(IsNull(oAtenciones("FechaEgreso")), Format(Now, "dd/mm/yyyy"), oAtenciones("FechaEgreso")) & " " & IIf(IsNull(oAtenciones("HoraEgreso")), Format(Now, "hh:mm:ss"), Format(oAtenciones("HoraEgreso"), "hh:mm:ss")) 'oAtenciones("FechaEgreso") & " " & Format(oAtenciones("HoraEgreso"), "hh:mm:ss")
  lcFuenteFinanciamientoElegida = "Fte.Financ: " & Left(oAtenciones!fuenteFinanciamiento, 15) & "  T.Serv: " & oAtenciones!TipoServicio
  Set grdServicios.DataSource = oServicios
  grdServicios.Enabled = True
End Sub

Private Sub grdAtenciones_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
  grdAtenciones.Bands(0).Columns("idCuentaAtencion").Header.Caption = "Cuenta Atención"
  grdAtenciones.Bands(0).Columns("idCuentaAtencion").Width = 1300
  grdAtenciones.Bands(0).Columns("idAtencion").Hidden = True
  'grdAtenciones.Bands(0).Columns("idAtencion").Width = 1000
  grdAtenciones.Bands(0).Columns("FechaIngreso").Header.Caption = "F.Ingreso"
  grdAtenciones.Bands(0).Columns("FechaIngreso").Width = 1200
  grdAtenciones.Bands(0).Columns("HoraIngreso").Header.Caption = "Hr.Ingreso"
  grdAtenciones.Bands(0).Columns("HoraIngreso").Width = 1000
  grdAtenciones.Bands(0).Columns("FechaEgreso").Header.Caption = "F.Egreso"
  grdAtenciones.Bands(0).Columns("FechaEgreso").Width = 1200
  grdAtenciones.Bands(0).Columns("HoraEgreso").Header.Caption = "Hr.Egreso"
  grdAtenciones.Bands(0).Columns("HoraEgreso").Width = 1000
  grdAtenciones.Bands(0).Columns("TipoServicio").Width = 2500
  grdAtenciones.Bands(0).Columns("idFormaPago").Hidden = True
  grdAtenciones.Bands(0).Columns("idTipoServicio").Hidden = True
  gridInfra.ConfigurarFilasBiColores grdAtenciones, sighentidades.GrillaConFilasBicolor
End Sub


Private Sub grdServicios_Click()
  Dim oConexion As New Connection
  oConexion.Open sighentidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  '
  Label15.Caption = "Servicio ..."
  btnAceptar.Enabled = False
  If oServicios.EOF = True And oServicios.BOF = True Then Exit Sub
  Frame1.Enabled = True
  btnAceptar.Enabled = True
  ml_idServicio = oServicios!idServicio
  If Not (IsNull(oServicios!IdCama)) Then
    ml_IdCama = oServicios!IdCama
  Else
    ml_IdCama = 0
  End If
  If Val(ml_IdCama) <> 0 Then
    Set oDOCama = mo_ReglasHoteleria.CamasSeleccionarPorId(ml_IdCama, oConexion)
    ml_Cama = oDOCama.codigo
  Else
    ml_Cama = ""
  End If
  txtNC.Text = ml_Cama
  ml_Servicio = oServicios!NombreServicio
  Label15.Caption = "Servicio: " & ml_Servicio
  If Not (IsDate(txtFechaInicio.Text)) Then txtFechaInicio.Text = IIf(IsNull(oAtenciones("FechaOcupacion")), Now, oAtenciones("FechaOcupacion")) & " " & IIf(IsNull(oAtenciones("HoraOcupacion")), Now, Format(oAtenciones("HoraOcupacion"), "hh:mm:ss"))
  If Not (IsDate(txtFechaFin.Text)) Then txtFechaFin.Text = IIf(IsNull(oAtenciones("FechaDesocupacion")), Format(Now, "dd/mm/yyyy"), oAtenciones("FechaDesocupacion")) & " " & IIf(IsNull(oAtenciones("HoraDesocupacion")), Format(Now, "hh:mm:ss"), Format(oAtenciones("HoraDesocupacion"), "hh:mm:ss"))
  If Val(ml_idServicio) <> 0 Then
    If optTodos.Value = True Then
      Set oLab = mo_ReglasLaboratorio.SeleccionaLaboratorioPorCuentaYServicio(ml_idCtaAtencion, ml_idServicio)
      Set oImag = mo_ReglasLaboratorio.SeleccionaImagenologiaPorCuentaYServicio(ml_idCtaAtencion, ml_idServicio)
      Set oIns = mo_ReglasLaboratorio.SeleccionaInsumosPorCuentaYServicio(ml_idCtaAtencion, ml_idServicio)
      Set oMed = mo_ReglasLaboratorio.SeleccionaFarmaciaPorCuentaYServicio(ml_idCtaAtencion, ml_idServicio)
      Set oProc = mo_ReglasLaboratorio.SeleccionaProcedimientosPorCuentaYServicio(ml_idCtaAtencion, ml_idServicio)
      Set oDevolucion = mo_ReglasLaboratorio.SeleccionaDevolucionesPorCuentaYServicio(ml_idCtaAtencion, ml_idServicio)
    ElseIf optFechas.Value = True Then
      Set oLab = mo_ReglasLaboratorio.SeleccionaLaboratorioPorCuentaYServicioYFecha(ml_idCtaAtencion, ml_idServicio, CDate(txtFechaInicio.Text), CDate(txtFechaFin.Text))
      Set oImag = mo_ReglasLaboratorio.SeleccionaImagenologiaPorCuentaYServicioYFecha(ml_idCtaAtencion, ml_idServicio, CDate(txtFechaInicio.Text), CDate(txtFechaFin.Text))
      Set oIns = mo_ReglasLaboratorio.SeleccionaInsumosPorCuentaYServicioYFecha(ml_idCtaAtencion, ml_idServicio, CDate(txtFechaInicio.Text), CDate(txtFechaFin.Text))
      Set oMed = mo_ReglasLaboratorio.SeleccionaFarmaciaPorCuentaYServicioYFecha(ml_idCtaAtencion, ml_idServicio, CDate(txtFechaInicio.Text), CDate(txtFechaFin.Text))
      Set oProc = mo_ReglasLaboratorio.SeleccionaProcedimientosPorCuentaYServicioYFecha(ml_idCtaAtencion, ml_idServicio, CDate(txtFechaInicio.Text), CDate(txtFechaFin.Text))
      Set oDevolucion = mo_ReglasLaboratorio.SeleccionadevolucionesPorCuentaYServicioYFecha(ml_idCtaAtencion, ml_idServicio, CDate(txtFechaInicio.Text), CDate(txtFechaFin.Text))
    Else
      Exit Sub
    End If
    Set ssLab.DataSource = oLab
    Set ssImag.DataSource = oImag
    Set ssIns.DataSource = oIns
    Set ssMed.DataSource = oMed
    Set ssProc.DataSource = oProc
  Else
    Set ssLab.DataSource = Nothing
    Set ssImag.DataSource = Nothing
    Set ssIns.DataSource = Nothing
    Set ssMed.DataSource = Nothing
    Set ssProc.DataSource = Nothing
  End If
  oConexion.Close
  Set oConexion = Nothing
End Sub

Private Sub grdServicios_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
  grdServicios.Bands(0).Columns("idEstanciaHospitalaria").Header.Caption = "Id Estancia"
  grdServicios.Bands(0).Columns("idEstanciaHospitalaria").Width = 900
  grdServicios.Bands(0).Columns("NombreServicio").Header.Caption = "Servicio"
  grdServicios.Bands(0).Columns("NombreServicio").Width = 4000
  grdServicios.Bands(0).Columns("IdServicio").Header.Caption = "Id Servicio"
  grdServicios.Bands(0).Columns("IdServicio").Width = 1000
  grdServicios.Bands(0).Columns("FechaOcupacion").Header.Caption = "Fecha Ingreso"
  grdServicios.Bands(0).Columns("FechaOcupacion").Width = 1300
  grdServicios.Bands(0).Columns("FechaOcupacion").Format = "dd/mm/yyyy"
  grdServicios.Bands(0).Columns("HoraOcupacion").Header.Caption = "Hora Ingreso"
  grdServicios.Bands(0).Columns("HoraOcupacion").Width = 1200
  grdServicios.Bands(0).Columns("HoraOcupacion").Format = "hh:mm:ss"
  grdServicios.Bands(0).Columns("FechaDesocupacion").Header.Caption = "Fecha Salida"
  grdServicios.Bands(0).Columns("FechaDesocupacion").Width = 1300
  grdServicios.Bands(0).Columns("FechaDesocupacion").Format = "dd/mm/yyyy"
  grdServicios.Bands(0).Columns("HoraDesocupacion").Header.Caption = "Hora Salida"
  grdServicios.Bands(0).Columns("HoraDesocupacion").Width = 1200
  grdServicios.Bands(0).Columns("HoraDesocupacion").Format = "hh:mm:ss"
  grdServicios.Bands(0).Columns("IdCama").Header.Caption = "Id Cama"
  grdServicios.Bands(0).Columns("IdCama").Width = 900
  gridInfra.ConfigurarFilasBiColores grdServicios, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub optTodos_Click(Value As Integer)
  If optTodos.Value = True Then
    Label4.Visible = False
    Label5.Visible = False
    txtFechaInicio.Visible = False
    txtFechaFin.Visible = False
  End If
End Sub

Private Sub optFechas_Click(Value As Integer)
  If optFechas.Value = True Then
    Label4.Visible = True
    Label5.Visible = True
    txtFechaInicio.Visible = True
    txtFechaFin.Visible = True
  End If
End Sub

Private Sub ssImag_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
  ssImag.Bands(0).Columns("tipo").Hidden = True
  ssImag.Bands(0).Columns("codigo").Width = 1000
  ssImag.Bands(0).Columns("codigo").Header.Caption = "Código"
  ssImag.Bands(0).Columns("Nombre").Width = 6500
  ssImag.Bands(0).Columns("cantidad").Width = 800
  ssImag.Bands(0).Columns("cantidad").Header.Caption = "Cantidad"
  ssImag.Bands(0).Columns("cantidad").Width = 800
  ssImag.Bands(0).Columns("cantidad").Header.Caption = "Cantidad"
  ssImag.Bands(0).Columns("precio").Width = 800
  ssImag.Bands(0).Columns("precio").Header.Caption = "Precio"
  ssImag.Bands(0).Columns("precio").Format = "0.00"
  ssImag.Bands(0).Columns("total").Width = 800
  ssImag.Bands(0).Columns("total").Header.Caption = "Total"
  ssImag.Bands(0).Columns("total").Format = "0.00"
  ssImag.Bands(0).Columns("idOrden").Hidden = True
  ssImag.Bands(0).Columns("idCuentaAtencion").Hidden = True
  ssImag.Bands(0).Columns("idTipoFinanciamiento").Hidden = True
  ssImag.Bands(0).Columns("FechaCreacion").Hidden = True
  ssImag.Bands(0).Columns("Descripcion").Hidden = True
  ssImag.Bands(0).Columns("Dfinanciamiento").Hidden = True
  ssImag.Bands(0).Columns("idEstadoFacturacion").Hidden = True
  ssImag.Bands(0).Columns("idFuenteFinanciamiento").Hidden = True
  ssImag.Bands(0).Columns("idPuntoCarga").Hidden = True
  ssImag.Bands(0).Columns("idServicioPaciente").Hidden = True
  ssImag.Bands(0).Columns("idProducto").Hidden = True
  gridInfra.ConfigurarFilasBiColores ssImag, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub ssIns_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
  ssIns.Bands(0).Columns("tipo").Hidden = True
  ssIns.Bands(0).Columns("codigo").Width = 1000
  ssIns.Bands(0).Columns("codigo").Header.Caption = "Código"
  ssIns.Bands(0).Columns("Nombre").Width = 6500
  ssIns.Bands(0).Columns("cantidad").Width = 800
  ssIns.Bands(0).Columns("cantidad").Header.Caption = "Cantidad"
  ssIns.Bands(0).Columns("precio").Width = 800
  ssIns.Bands(0).Columns("precio").Header.Caption = "Precio"
  ssIns.Bands(0).Columns("precio").Format = "0.00"
  ssIns.Bands(0).Columns("total").Width = 800
  ssIns.Bands(0).Columns("total").Header.Caption = "Total"
  ssIns.Bands(0).Columns("total").Format = "0.00"
  ssIns.Bands(0).Columns("idFuenteFinanciamiento").Hidden = True
  ssIns.Bands(0).Columns("idTipoFinanciamiento").Hidden = True
  ssIns.Bands(0).Columns("idTipoFinanAtenciones").Hidden = True
  ssIns.Bands(0).Columns("idServicioPaciente").Hidden = True
  ssIns.Bands(0).Columns("idProducto").Hidden = True
  gridInfra.ConfigurarFilasBiColores ssIns, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub ssLab_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
  ssLab.Bands(0).Columns("tipo").Hidden = True
  ssLab.Bands(0).Columns("codigo").Width = 1000
  ssLab.Bands(0).Columns("codigo").Header.Caption = "Código"
  ssLab.Bands(0).Columns("Nombre").Width = 6500
  ssLab.Bands(0).Columns("cantidad").Width = 800
  ssLab.Bands(0).Columns("cantidad").Header.Caption = "Cantidad"
  ssLab.Bands(0).Columns("cantidad").Width = 800
  ssLab.Bands(0).Columns("cantidad").Header.Caption = "Cantidad"
  ssLab.Bands(0).Columns("precio").Width = 800
  ssLab.Bands(0).Columns("precio").Header.Caption = "Precio"
  ssLab.Bands(0).Columns("precio").Format = "0.00"
  ssLab.Bands(0).Columns("total").Width = 800
  ssLab.Bands(0).Columns("total").Header.Caption = "Total"
  ssLab.Bands(0).Columns("total").Format = "0.00"
  ssLab.Bands(0).Columns("idOrden").Hidden = True
  ssLab.Bands(0).Columns("idCuentaAtencion").Hidden = True
  ssLab.Bands(0).Columns("idTipoFinanciamiento").Hidden = True
  ssLab.Bands(0).Columns("FechaCreacion").Hidden = True
  ssLab.Bands(0).Columns("Descripcion").Hidden = True
  ssLab.Bands(0).Columns("Dfinanciamiento").Hidden = True
  ssLab.Bands(0).Columns("idEstadoFacturacion").Hidden = True
  ssLab.Bands(0).Columns("idFuenteFinanciamiento").Hidden = True
  ssLab.Bands(0).Columns("idPuntoCarga").Hidden = True
  ssLab.Bands(0).Columns("idServicioPaciente").Hidden = True
  ssLab.Bands(0).Columns("idProducto").Hidden = True
  gridInfra.ConfigurarFilasBiColores ssLab, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub ssMed_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
  ssMed.Bands(0).Columns("tipo").Hidden = True
  ssMed.Bands(0).Columns("codigo").Width = 1000
  ssMed.Bands(0).Columns("codigo").Header.Caption = "Código"
  ssMed.Bands(0).Columns("Nombre").Width = 6500
  ssMed.Bands(0).Columns("cantidad").Width = 800
  ssMed.Bands(0).Columns("cantidad").Header.Caption = "Cantidad"
  ssMed.Bands(0).Columns("precio").Width = 800
  ssMed.Bands(0).Columns("precio").Header.Caption = "Precio"
  ssMed.Bands(0).Columns("precio").Format = "0.00"
  ssMed.Bands(0).Columns("total").Width = 800
  ssMed.Bands(0).Columns("total").Header.Caption = "Total"
  ssMed.Bands(0).Columns("total").Format = "0.00"
  ssMed.Bands(0).Columns("idFuenteFinanciamiento").Hidden = True
  ssMed.Bands(0).Columns("idTipoFinanciamiento").Hidden = True
  ssMed.Bands(0).Columns("idTipoFinanAtenciones").Hidden = True
  ssMed.Bands(0).Columns("idServicioPaciente").Hidden = True
  ssMed.Bands(0).Columns("idProducto").Hidden = True
  gridInfra.ConfigurarFilasBiColores ssMed, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub ssProc_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
  ssProc.Bands(0).Columns("tipo").Hidden = True
  ssProc.Bands(0).Columns("codigo").Width = 1000
  ssProc.Bands(0).Columns("codigo").Header.Caption = "Código"
  ssProc.Bands(0).Columns("Nombre").Width = 6500
  ssProc.Bands(0).Columns("cantidad").Width = 800
  ssProc.Bands(0).Columns("cantidad").Header.Caption = "Cantidad"
  ssProc.Bands(0).Columns("cantidad").Width = 800
  ssProc.Bands(0).Columns("cantidad").Header.Caption = "Cantidad"
  ssProc.Bands(0).Columns("precio").Width = 800
  ssProc.Bands(0).Columns("precio").Header.Caption = "Precio"
  ssProc.Bands(0).Columns("precio").Format = "0.00"
  ssProc.Bands(0).Columns("total").Width = 800
  ssProc.Bands(0).Columns("total").Header.Caption = "Total"
  ssProc.Bands(0).Columns("total").Format = "0.00"
  ssProc.Bands(0).Columns("idOrden").Hidden = True
  ssProc.Bands(0).Columns("idCuentaAtencion").Hidden = True
  ssProc.Bands(0).Columns("idTipoFinanciamiento").Hidden = True
  ssProc.Bands(0).Columns("FechaCreacion").Hidden = True
  ssProc.Bands(0).Columns("Descripcion").Hidden = True
  ssProc.Bands(0).Columns("Dfinanciamiento").Hidden = True
  ssProc.Bands(0).Columns("idEstadoFacturacion").Hidden = True
  ssProc.Bands(0).Columns("idFuenteFinanciamiento").Hidden = True
  ssProc.Bands(0).Columns("idPuntoCarga").Hidden = True
  ssProc.Bands(0).Columns("idServicioPaciente").Hidden = True
  ssProc.Bands(0).Columns("idProducto").Hidden = True
  gridInfra.ConfigurarFilasBiColores ssProc, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub txtCA_LostFocus()
  ml_CAfiliacion = Trim(txtCA.Text)
End Sub

Private Sub txtFechaFin_Change()
  Set ssLab.DataSource = Nothing
  Set ssImag.DataSource = Nothing
  Set ssIns.DataSource = Nothing
  Set ssMed.DataSource = Nothing
  Set ssProc.DataSource = Nothing
End Sub

Private Sub txtFechaFin_LostFocus()
    If txtFechaFin <> sighentidades.FECHA_VACIA_DMY_HMS Then
        If Not IsDate(txtFechaFin) Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFechaFin = sighentidades.FECHA_VACIA_DMY_HMS
        End If
    End If
End Sub

Private Sub txtFechaInicio_Change()
  Set ssLab.DataSource = Nothing
  Set ssImag.DataSource = Nothing
  Set ssIns.DataSource = Nothing
  Set ssMed.DataSource = Nothing
  Set ssProc.DataSource = Nothing
End Sub

Private Sub txtFechaInicio_LostFocus()
    If txtFechaInicio <> sighentidades.FECHA_VACIA_DMY_HMS Then
        If Not IsDate(txtFechaInicio) Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFechaInicio = sighentidades.FECHA_VACIA_DMY_HMS
        End If
    End If
End Sub

Private Sub txtHC_KeyPress(KeyAscii As Integer)
  If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 22 Or KeyAscii = 3) Then KeyAscii = 0
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtHC_LostFocus()
  If Trim(txtHC.Text) = "" Then
    cmdBuscaCuentaPorApellidos_Click
    Exit Sub
  End If
  Set grdAtenciones.DataSource = Nothing
  Set grdServicios.DataSource = Nothing
  Set ssImag.DataSource = Nothing
  Set ssIns.DataSource = Nothing
  Set ssLab.DataSource = Nothing
  Set ssMed.DataSource = Nothing
  Set ssProc.DataSource = Nothing
  txtAN.Text = ""
  txtNC.Text = ""
  txtD.Text = ""
  txtFA.Text = ""
  grdServicios.Enabled = False
  Frame1.Enabled = False
  ml_Historia = Val(txtHC.Text)
  BuscaPaciente ml_Historia
End Sub

Private Function AveriguaDevueltos1(CodProducto) As Long
  AveriguaDevueltos1 = 0
  If oDevolucion.State = adStateClosed Then Exit Function
  If oDevolucion.EOF = True And oDevolucion.BOF = True Then Exit Function
  oDevolucion.MoveFirst
  Do While Not oDevolucion.EOF
    If oDevolucion!codigo = CodProducto Then AveriguaDevueltos1 = AveriguaDevueltos1 + oDevolucion!Cantidad
    oDevolucion.MoveNext
  Loop
  AveriguaDevueltos1 = AveriguaDevueltos1 / 3
End Function

