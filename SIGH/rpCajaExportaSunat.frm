VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form rpCajaExportaSunat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generador de Tramas - Facturador Sunat"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11490
   Icon            =   "rpCajaExportaSunat.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   11490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tabFacturadorSunat 
      Height          =   7800
      Left            =   105
      TabIndex        =   0
      Top             =   90
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   13758
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Exporta Tramas Facturador Sunat"
      TabPicture(0)   =   "rpCajaExportaSunat.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grdResumenDiario"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Configuración Notas de Credito"
      TabPicture(1)   =   "rpCajaExportaSunat.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame(0)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Carpetas Facturador Sunat"
      TabPicture(2)   =   "rpCajaExportaSunat.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(1)=   "Frame(1)"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame 
         Height          =   4710
         Index           =   2
         Left            =   8805
         TabIndex        =   49
         Top             =   1920
         Visible         =   0   'False
         Width           =   2355
         Begin VB.Frame Frame 
            Caption         =   "Ticket Sunat"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1500
            Index           =   3
            Left            =   105
            TabIndex        =   52
            Top             =   1560
            Visible         =   0   'False
            Width           =   2205
            Begin VB.CommandButton cmdGrabaTicket 
               Caption         =   "Graba Ticket SUNAT"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   720
               Left            =   105
               TabIndex        =   54
               Top             =   645
               Width           =   2025
            End
            Begin VB.TextBox txtTicketSunat 
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
               Left            =   75
               TabIndex        =   53
               Top             =   270
               Width           =   2070
            End
         End
         Begin VB.CommandButton cmdEliminaRD 
            Caption         =   "Elimina RESUMEN DIARIO ?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   120
            TabIndex        =   51
            Top             =   3840
            Visible         =   0   'False
            Width           =   2160
         End
         Begin VB.CommandButton cmdSoloMuestra 
            Caption         =   "Solo muestra RESUMEN DIARIO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   120
            TabIndex        =   50
            Top             =   150
            Visible         =   0   'False
            Width           =   2160
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Configuración Notas de Credito"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Index           =   0
         Left            =   -74925
         TabIndex        =   14
         Top             =   360
         Width           =   11115
         Begin VB.Frame Frame8 
            Caption         =   "FACTURA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1995
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   5385
            Begin VB.TextBox txtNumInicialF 
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
               Left            =   1680
               MaxLength       =   8
               TabIndex        =   40
               Top             =   735
               Width           =   1785
            End
            Begin VB.TextBox txtNumFinalF 
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
               Left            =   1680
               MaxLength       =   8
               TabIndex        =   39
               Top             =   1095
               Width           =   1785
            End
            Begin VB.TextBox txtNumeroUltimoF 
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
               Left            =   1680
               MaxLength       =   8
               TabIndex        =   38
               Top             =   1455
               Width           =   1785
            End
            Begin VB.TextBox txtNroSerieNotaF 
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
               Left            =   1680
               MaxLength       =   4
               TabIndex        =   37
               Top             =   375
               Width           =   1785
            End
            Begin VB.TextBox txtNroSerieNotaS 
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
               Left            =   3495
               MaxLength       =   4
               TabIndex        =   36
               Top             =   375
               Width           =   1785
            End
            Begin VB.TextBox txtNumeroUltimoS 
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
               Left            =   3495
               MaxLength       =   8
               TabIndex        =   35
               Top             =   1455
               Width           =   1785
            End
            Begin VB.TextBox txtNumFinalS 
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
               Left            =   3495
               MaxLength       =   8
               TabIndex        =   34
               Top             =   1095
               Width           =   1785
            End
            Begin VB.TextBox txtNumInicialS 
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
               Left            =   3495
               MaxLength       =   8
               TabIndex        =   33
               Top             =   735
               Width           =   1785
            End
            Begin VB.Label Label11 
               Caption         =   "Rango Inicial"
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
               TabIndex        =   45
               Top             =   810
               Width           =   1215
            End
            Begin VB.Label Label12 
               Caption         =   "Rango Final"
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
               TabIndex        =   44
               Top             =   1200
               Width           =   1365
            End
            Begin VB.Label Label13 
               Caption         =   "último generado"
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
               TabIndex        =   43
               Top             =   1530
               Width           =   1335
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Nro Serie"
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
               TabIndex        =   42
               Top             =   360
               Width           =   750
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Farmacia                   Servicios"
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
               Left            =   2265
               TabIndex        =   41
               Top             =   150
               Width           =   2520
            End
         End
         Begin VB.Frame Frame2 
            Height          =   1080
            Left            =   120
            TabIndex        =   29
            Top             =   2400
            Width           =   10905
            Begin VB.CommandButton btnGrabar 
               Caption         =   "Grabar"
               DisabledPicture =   "rpCajaExportaSunat.frx":0D1E
               DownPicture     =   "rpCajaExportaSunat.frx":117E
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
               Left            =   4414
               Picture         =   "rpCajaExportaSunat.frx":15F3
               Style           =   1  'Graphical
               TabIndex        =   31
               Top             =   225
               Width           =   1365
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Cancelar (ESC)"
               DisabledPicture =   "rpCajaExportaSunat.frx":1A68
               DownPicture     =   "rpCajaExportaSunat.frx":1F2C
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
               Left            =   5906
               Picture         =   "rpCajaExportaSunat.frx":2418
               Style           =   1  'Graphical
               TabIndex        =   30
               Top             =   225
               Width           =   1365
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "BOLETA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1995
            Left            =   5580
            TabIndex        =   15
            Top             =   360
            Width           =   5430
            Begin VB.TextBox txtNroSerieNotaB 
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
               Left            =   1680
               MaxLength       =   4
               TabIndex        =   23
               Top             =   360
               Width           =   1785
            End
            Begin VB.TextBox txtNumeroUltimoB 
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
               Left            =   1680
               MaxLength       =   8
               TabIndex        =   22
               Top             =   1440
               Width           =   1785
            End
            Begin VB.TextBox txtNumFinalB 
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
               Left            =   1680
               MaxLength       =   8
               TabIndex        =   21
               Top             =   1080
               Width           =   1785
            End
            Begin VB.TextBox txtNumInicialB 
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
               Left            =   1680
               MaxLength       =   8
               TabIndex        =   20
               Top             =   720
               Width           =   1785
            End
            Begin VB.TextBox txtNroSerieNotaBS 
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
               Left            =   3480
               MaxLength       =   4
               TabIndex        =   19
               Top             =   345
               Width           =   1785
            End
            Begin VB.TextBox txtNumeroUltimoBS 
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
               Left            =   3480
               MaxLength       =   8
               TabIndex        =   18
               Top             =   1425
               Width           =   1785
            End
            Begin VB.TextBox txtNumFinalBS 
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
               Left            =   3480
               MaxLength       =   8
               TabIndex        =   17
               Top             =   1065
               Width           =   1785
            End
            Begin VB.TextBox txtNumInicialBS 
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
               Left            =   3480
               MaxLength       =   8
               TabIndex        =   16
               Top             =   705
               Width           =   1785
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Nro Serie"
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
               Top             =   360
               Width           =   750
            End
            Begin VB.Label Label4 
               Caption         =   "último generado"
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
               TabIndex        =   27
               Top             =   1530
               Width           =   1335
            End
            Begin VB.Label Label5 
               Caption         =   "Rango Final"
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
               TabIndex        =   26
               Top             =   1200
               Width           =   1365
            End
            Begin VB.Label Label6 
               Caption         =   "Rango Inicial"
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
               TabIndex        =   25
               Top             =   810
               Width           =   1215
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Farmacia                   Servicios"
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
               Left            =   2265
               TabIndex        =   24
               Top             =   135
               Width           =   2520
            End
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1080
         Left            =   -74880
         TabIndex        =   11
         Top             =   2160
         Width           =   11040
         Begin VB.CommandButton btnGuardarConfFacturador 
            Caption         =   "Grabar"
            DisabledPicture =   "rpCajaExportaSunat.frx":2904
            DownPicture     =   "rpCajaExportaSunat.frx":2D64
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
            Left            =   4317
            Picture         =   "rpCajaExportaSunat.frx":31D9
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   225
            Width           =   1365
         End
         Begin VB.CommandButton Command2 
            Cancel          =   -1  'True
            Caption         =   "Cancelar (ESC)"
            DisabledPicture =   "rpCajaExportaSunat.frx":364E
            DownPicture     =   "rpCajaExportaSunat.frx":3B12
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
            Left            =   5809
            Picture         =   "rpCajaExportaSunat.frx":3FFE
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   240
            Width           =   1365
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "FACTURADOR SUNAT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Index           =   1
         Left            =   -74880
         TabIndex        =   8
         Top             =   360
         Width           =   11070
         Begin VB.TextBox txtDataFacturador 
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
            Left            =   720
            TabIndex        =   9
            Top             =   360
            Width           =   10035
         End
         Begin VB.Label Label7 
            Caption         =   "DATA"
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
            TabIndex        =   10
            Top             =   360
            Width           =   585
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1080
         Left            =   120
         TabIndex        =   6
         Top             =   6585
         Width           =   11040
         Begin VB.CommandButton btnAceptar 
            Caption         =   "Exportar "
            DisabledPicture =   "rpCajaExportaSunat.frx":44EA
            DownPicture     =   "rpCajaExportaSunat.frx":494A
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
            Left            =   4247
            Picture         =   "rpCajaExportaSunat.frx":4DBF
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   225
            Width           =   1365
         End
         Begin VB.CommandButton btnCancelar 
            Caption         =   "Cancelar (ESC)"
            DisabledPicture =   "rpCajaExportaSunat.frx":5234
            DownPicture     =   "rpCajaExportaSunat.frx":56F8
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
            Left            =   5809
            Picture         =   "rpCajaExportaSunat.frx":5BE4
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   225
            Width           =   1365
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Exporta Tramas Facturador Sunat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1560
         Left            =   120
         TabIndex        =   1
         Top             =   375
         Width           =   11040
         Begin VB.CheckBox chkUsaResumenDiario 
            Alignment       =   1  'Right Justify
            Caption         =   "Usa RESUMEN DIARIO SUNAT"
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
            Left            =   105
            TabIndex        =   46
            Top             =   1155
            Width           =   3030
         End
         Begin MSMask.MaskEdBox txtFechaInicio 
            Height          =   315
            Left            =   2925
            TabIndex        =   2
            Top             =   405
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
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
         Begin MSMask.MaskEdBox txtFechaFin 
            Height          =   315
            Left            =   2925
            TabIndex        =   3
            Top             =   795
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
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
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "F.Emisión Documento Final"
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
            TabIndex        =   5
            Top             =   855
            Width           =   2175
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "F.Emisión Documento Inicial"
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
            TabIndex        =   4
            Top             =   450
            Width           =   2265
         End
      End
      Begin UltraGrid.SSUltraGrid grdResumenDiario 
         Height          =   4605
         Left            =   120
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   1995
         Visible         =   0   'False
         Width           =   8670
         _ExtentX        =   15293
         _ExtentY        =   8123
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
         Caption         =   "Resumen Diario"
      End
   End
End
Attribute VB_Name = "rpCajaExportaSunat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reporte de exportacion sunat
'        Programado por: Cachay F
'        Fecha: Enero 2016
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Apariencia As New sighEntidades.GridInfragistic
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_cmbCentroCostos As New ListaDespleglable
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim ml_IdTipoReporte As Long
Dim ml_idUsuario As Long
Dim mo_lcNombrePc As String
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_sighProxies As New SIGHProxies.Procesos
Dim oRsTmpDctos As New Recordset
Dim oDONotaCreditoDebitoTipoNotaF As New DONotaCreditoDebitoTipoNota
Dim oDONotaCreditoDebitoTipoNotaB As New DONotaCreditoDebitoTipoNota
Dim oDONotaCreditoDebitoTipoNotaFS As New DONotaCreditoDebitoTipoNota
Dim oDONotaCreditoDebitoTipoNotaBS As New DONotaCreditoDebitoTipoNota
Dim ldHoy As Date, lcFechaEmision As String, lbNoPulsoBuscarHistorico As Boolean, lbTieneTicketSunat As Boolean
Const TipoNota As Integer = 2
Const TipoFactura As Long = 2
Const TipoBoleta As Long = 3
Const TipoNotaServicio As Integer = 3

Property Let idUsuario(lIdValue As Long)
    ml_idUsuario = lIdValue
End Property
Property Let IdTipoReporte(lIdValue As Long)
    ml_IdTipoReporte = lIdValue
End Property
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property

Private Sub btnAceptar_Click()
    If btnAceptar.Visible = False Then
       Exit Sub
    End If
    
    If IsDate(txtFechaInicio.Text) = False Then
       MsgBox "La FECHA INICIAL es vacia ó no tiene el formato correcto", vbInformation, "Exporta Facturador Sunat"
       Exit Sub
    End If
    If IsDate(txtFechaFin.Text) = False Then
       MsgBox "La FECHA FINAL es vacia ó no tiene el formato correcto", vbInformation, "Exporta Facturador Sunat"
       Exit Sub
    End If
    If CDate(txtFechaInicio.Text) > CDate(txtFechaFin.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, "Exporta Facturador Sunat"
       Exit Sub
    End If
    If chkUsaResumenDiario.Value = 1 Then
       If CDate(Format(Me.txtFechaInicio.Text, sighEntidades.DevuelveFechaSoloFormato_DMY)) <> CDate(Format(Me.txtFechaFin.Text, sighEntidades.DevuelveFechaSoloFormato_DMY)) Then
            MsgBox "La FECHA FINAL debe ser igual a la FECHA INICIAL", vbInformation, "Exporta Facturador Sunat"
            Exit Sub
       ElseIf CDate(Me.txtFechaInicio.Text) >= ldHoy Then
            MsgBox "La FECHA INICIAL debe ser menor a HOY (Se emite PARTE DIARIO solo de días pasados)", vbInformation, "Exporta Facturador Sunat"
            Exit Sub
       End If
       If lbNoPulsoBuscarHistorico = False Then
            MsgBox "Antes de generar archivos planos, tiene que ver si hay un Histórico" & Chr(13) & _
                   "      pulsando clic en botón 'SOLO MUESTRA RESUMEN DIARIO'", vbInformation, "Exporta Facturador Sunat"
            Exit Sub
       End If
    End If
    MousePointer = 11
    Dim oExportar As New SIGHProxies.Procesos
    Dim oRsBusquedaRecibos As New Recordset
    Dim oRsAnuladosComoActivos As New Recordset
    
    Dim lbContinuar As Boolean, lbYaTieneHistoricoDeRD As Boolean
    lbContinuar = True
    lbYaTieneHistoricoDeRD = False
    grdResumenDiario.Visible = False
    If chkUsaResumenDiario.Value = 1 Then
         If lbTieneTicketSunat = True Then
            lbContinuar = False
            lbYaTieneHistoricoDeRD = True
         ElseIf Me.cmdEliminaRD.Visible = True Then
            lbContinuar = False
            lbYaTieneHistoricoDeRD = True
         End If
'        lcFechaEmision = Year(CDate(Me.txtFechaInicio.Text)) & "-" & Right("00" & Month(CDate(Me.txtFechaInicio.Text)), 2) & "-" & _
'                         Right("00" & Day(CDate(Me.txtFechaInicio.Text)), 2)
'        Set oRsTmpDctos = mo_AdminCaja.Sunat_ResumenDiarioSeleccionarPorFechaEmision(lcFechaEmision)
'        If oRsTmpDctos.RecordCount > 0 Then
'           If Not IsNull(oRsTmpDctos!DctoRDI) Then
'              If Len(Trim(oRsTmpDctos!DctoRDI)) = 27 Then
'                  lbYaTieneHistoricoDeRD = True
'                  lbContinuar = False
'              End If
'           End If
'        End If
'        grdResumenDiario.Visible = True
'        Set grdResumenDiario.DataSource = oRsTmpDctos
'        mo_Apariencia.ConfigurarFilasBiColores grdResumenDiario, sighentidades.GrillaConFilasBicolor
    End If
    If lbContinuar = True Then
        'procesar las ANULADAS como ACTIVAS
        Set oRsBusquedaRecibos = mo_AdminCaja.CajaComprobantePagoFiltroPorNroSerieDocumentoOporRangoFemision("", "", CDate(txtFechaInicio.Text), CDate(txtFechaFin.Text))
        oRsBusquedaRecibos.Filter = "IdEstadoComprobante = 9"
        If oRsBusquedaRecibos.RecordCount > 0 Then
           With oRsAnuladosComoActivos
              .Fields.Append "Caja", adVarChar, 200, adFldIsNullable
              .Fields.Append "Turno", adVarChar, 200, adFldIsNullable
              .Fields.Append "Fecha", adDate
              .Fields.Append "NroSerie", adVarChar, 4, adFldIsNullable
              .Fields.Append "NroDocumento", adVarChar, 8, adFldIsNullable
              .Fields.Append "NroHistoriaClinica", adInteger
              .Fields.Append "RazonSocial", adVarChar, 200, adFldIsNullable
              .Fields.Append "total", adDouble
              .Fields.Append "IdCuentaAtencion", adInteger
              .Fields.Append "Estado", adVarChar, 200, adFldIsNullable
              .Fields.Append "CajeroApPat", adVarChar, 40, adFldIsNullable
              .Fields.Append "CajeroApMat", adVarChar, 40, adFldIsNullable
              .Fields.Append "CajeroNombres", adVarChar, 80, adFldIsNullable
              .Fields.Append "BienServicio", adVarChar, 200, adFldIsNullable
              .Fields.Append "IdPaciente", adInteger
              .Fields.Append "idTurno", adInteger
              .Fields.Append "idCaja", adInteger
              .Fields.Append "IdCajero", adInteger
              .Fields.Append "IdEmpleado", adInteger
              .Fields.Append "idFormaPago", adInteger
              .Fields.Append "dFormaPago", adVarChar, 200, adFldIsNullable
              .Fields.Append "IdEstadoComprobante", adInteger
              .Fields.Append "IdTipoOrden", adInteger
              .Fields.Append "idFarmacia", adInteger
              .Fields.Append "dFarmacia", adVarChar, 200, adFldIsNullable
              .Fields.Append "exoneraciones", adDouble
              .Fields.Append "IdTipoComprobante", adInteger
              .Fields.Append "SubTotal", adDouble
              .Fields.Append "FechaCobranza", adDate
              .Fields.Append "IGV", adDouble
              .Fields.Append "ruc", adVarChar, 20, adFldIsNullable
              .Fields.Append "IdComprobantePago", adInteger
              .Fields.Append "dni", adVarChar, 20, adFldIsNullable
              .LockType = adLockOptimistic
              .Open
           End With
           oRsBusquedaRecibos.MoveFirst
           Do While Not oRsBusquedaRecibos.EOF
              oRsAnuladosComoActivos.AddNew
              oRsAnuladosComoActivos!caja = IIf(IsNull(oRsBusquedaRecibos!caja), 0, oRsBusquedaRecibos!caja)
              oRsAnuladosComoActivos!Turno = IIf(IsNull(oRsBusquedaRecibos!Turno), "", oRsBusquedaRecibos!Turno)
              oRsAnuladosComoActivos!fecha = IIf(IsNull(oRsBusquedaRecibos!fecha), 0, oRsBusquedaRecibos!fecha)
              oRsAnuladosComoActivos!nroSerie = IIf(IsNull(oRsBusquedaRecibos!nroSerie), "", oRsBusquedaRecibos!nroSerie)
              oRsAnuladosComoActivos!nrodocumento = IIf(IsNull(oRsBusquedaRecibos!nrodocumento), "", oRsBusquedaRecibos!nrodocumento)
              oRsAnuladosComoActivos!NroHistoriaClinica = IIf(IsNull(oRsBusquedaRecibos!NroHistoriaClinica), 0, oRsBusquedaRecibos!NroHistoriaClinica)
              oRsAnuladosComoActivos!razonSocial = IIf(IsNull(oRsBusquedaRecibos!razonSocial), "", oRsBusquedaRecibos!razonSocial)
              oRsAnuladosComoActivos!Total = IIf(IsNull(oRsBusquedaRecibos!Total), 0, oRsBusquedaRecibos!Total)
              oRsAnuladosComoActivos!idCuentaAtencion = IIf(IsNull(oRsBusquedaRecibos!idCuentaAtencion), 0, oRsBusquedaRecibos!idCuentaAtencion)
              oRsAnuladosComoActivos!estado = IIf(IsNull(oRsBusquedaRecibos!estado), "", oRsBusquedaRecibos!estado)
              oRsAnuladosComoActivos!CajeroApPat = IIf(IsNull(oRsBusquedaRecibos!CajeroApPat), "", oRsBusquedaRecibos!CajeroApPat)
              oRsAnuladosComoActivos!CajeroApMat = IIf(IsNull(oRsBusquedaRecibos!CajeroApMat), "", oRsBusquedaRecibos!CajeroApMat)
              oRsAnuladosComoActivos!CajeroNombres = IIf(IsNull(oRsBusquedaRecibos!CajeroNombres), "", oRsBusquedaRecibos!CajeroNombres)
              oRsAnuladosComoActivos!BienServicio = IIf(IsNull(oRsBusquedaRecibos!BienServicio), "", oRsBusquedaRecibos!BienServicio)
              oRsAnuladosComoActivos!idPaciente = IIf(IsNull(oRsBusquedaRecibos!idPaciente), 0, oRsBusquedaRecibos!idPaciente)
              oRsAnuladosComoActivos!IdTurno = IIf(IsNull(oRsBusquedaRecibos!IdTurno), 0, oRsBusquedaRecibos!IdTurno)
              oRsAnuladosComoActivos!IdCaja = IIf(IsNull(oRsBusquedaRecibos!IdCaja), 0, oRsBusquedaRecibos!IdCaja)
              oRsAnuladosComoActivos!IdCajero = IIf(IsNull(oRsBusquedaRecibos!IdCajero), 0, oRsBusquedaRecibos!IdCajero)
              oRsAnuladosComoActivos!IdEmpleado = IIf(IsNull(oRsBusquedaRecibos!IdEmpleado), 0, oRsBusquedaRecibos!IdEmpleado)
              oRsAnuladosComoActivos!IdFormaPago = IIf(IsNull(oRsBusquedaRecibos!IdFormaPago), 0, oRsBusquedaRecibos!IdFormaPago)
              oRsAnuladosComoActivos!dFormaPago = IIf(IsNull(oRsBusquedaRecibos!dFormaPago), "", oRsBusquedaRecibos!dFormaPago)
              oRsAnuladosComoActivos!idEstadoComprobante = 4      'IIf(IsNull(oRsBusquedaRecibos!IdEstadoComprobante), 0, oRsBusquedaRecibos!IdEstadoComprobante)
              oRsAnuladosComoActivos!IdTipoOrden = IIf(IsNull(oRsBusquedaRecibos!IdTipoOrden), 0, oRsBusquedaRecibos!IdTipoOrden)
              oRsAnuladosComoActivos!idFarmacia = IIf(IsNull(oRsBusquedaRecibos!idFarmacia), 0, oRsBusquedaRecibos!idFarmacia)
              oRsAnuladosComoActivos!dFarmacia = IIf(IsNull(oRsBusquedaRecibos!dFarmacia), "", oRsBusquedaRecibos!dFarmacia)
              oRsAnuladosComoActivos!exoneraciones = IIf(IsNull(oRsBusquedaRecibos!exoneraciones), 0, oRsBusquedaRecibos!exoneraciones)
              oRsAnuladosComoActivos!IdTipoComprobante = IIf(IsNull(oRsBusquedaRecibos!IdTipoComprobante), 0, oRsBusquedaRecibos!IdTipoComprobante)
              oRsAnuladosComoActivos!Subtotal = IIf(IsNull(oRsBusquedaRecibos!Subtotal), 0, oRsBusquedaRecibos!Subtotal)
              oRsAnuladosComoActivos!FechaCobranza = IIf(IsNull(oRsBusquedaRecibos!FechaCobranza), 0, oRsBusquedaRecibos!FechaCobranza)
              oRsAnuladosComoActivos!IGV = IIf(IsNull(oRsBusquedaRecibos!IGV), 0, oRsBusquedaRecibos!IGV)
              oRsAnuladosComoActivos!ruc = IIf(IsNull(oRsBusquedaRecibos!ruc), "", oRsBusquedaRecibos!ruc)
              oRsAnuladosComoActivos!IdComprobantePago = IIf(IsNull(oRsBusquedaRecibos!IdComprobantePago), 0, oRsBusquedaRecibos!IdComprobantePago)
              oRsAnuladosComoActivos!DNI = IIf(IsNull(oRsBusquedaRecibos!DNI), "", oRsBusquedaRecibos!DNI)
              oRsAnuladosComoActivos.Update
              oRsBusquedaRecibos.MoveNext
           Loop
           oExportar.ExportarFacturasBoletas "", "", "", "", IIf(Me.chkUsaResumenDiario.Value = 1, True, False), oRsAnuladosComoActivos
        End If
        oRsBusquedaRecibos.Close
        '
        oExportar.ExportarFacturasBoletas txtFechaInicio.Text, txtFechaFin.Text, "", "", IIf(Me.chkUsaResumenDiario.Value = 1, True, False)
        oExportar.ExportarNotasCredito txtFechaInicio.Text, txtFechaFin.Text, "", "", IIf(Me.chkUsaResumenDiario.Value = 1, True, False)
        '
    End If
    If Me.chkUsaResumenDiario.Value = 1 Then
      '  If lbYaTieneHistoricoDeRD = False Then
           Set oRsTmpDctos = mo_AdminCaja.Sunat_ResumenDiarioSeleccionarPorFechaEmision(lcFechaEmision)
      '  End If
        oExportar.ExportarResumenDiarioYanulaciones12 lbYaTieneHistoricoDeRD, txtFechaInicio.Text, oRsAnuladosComoActivos, _
                                                      oRsTmpDctos
        MsgBox "Se exportó correctamente la trama de los comprobantes de pago (Facturas/Nota de Créditos Facturas/Parte Diario/Anulaciones)" & _
     Chr(13) & "                                         N° Dctos del Parte Diario: " & Trim(Str(oRsTmpDctos.RecordCount)), vbInformation, "Facturador Sunat (Exporta)"
    Else
        MsgBox "Se exportó correctamente la trama de los comprobantes de pago (Boletas/Facturas/Nota de Créditos)", vbInformation, "Facturador Sunat (Exporta)"
    End If
    Set oExportar = Nothing
    Me.Visible = False
    MousePointer = 1
End Sub



Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub

Sub CargaDatosAlObjetosDeDatos()
    With oDONotaCreditoDebitoTipoNotaF
        .idTipoNota = 2
        .nrodocumento = Me.txtNumeroUltimoF.Text
        .NroDocumentoFinal = Me.txtNumFinalF.Text
        .NroDocumentoInicial = Me.txtNumInicialF.Text
        .nroSerie = Me.txtNroSerieNotaF.Text
        .TipoNota = TipoFactura
        .IdUsuarioAuditoria = ml_idUsuario
    End With
    
    With oDONotaCreditoDebitoTipoNotaB
        .idTipoNota = 2
        .nrodocumento = Me.txtNumeroUltimoB.Text
        .NroDocumentoFinal = Me.txtNumFinalB.Text
        .NroDocumentoInicial = Me.txtNumInicialB.Text
        .nroSerie = Me.txtNroSerieNotaB.Text
        .TipoNota = TipoBoleta
        .IdUsuarioAuditoria = ml_idUsuario
    End With
    With oDONotaCreditoDebitoTipoNotaFS
        .idTipoNota = 3
        .nrodocumento = Me.txtNumeroUltimoS.Text
        .NroDocumentoFinal = Me.txtNumFinalS.Text
        .NroDocumentoInicial = Me.txtNumInicialS.Text
        .nroSerie = Me.txtNroSerieNotaS.Text
        .TipoNota = TipoFactura
        .IdUsuarioAuditoria = ml_idUsuario
    End With
    
    With oDONotaCreditoDebitoTipoNotaBS
        .idTipoNota = 3
        .nrodocumento = Me.txtNumeroUltimoBS.Text
        .NroDocumentoFinal = Me.txtNumFinalBS.Text
        .NroDocumentoInicial = Me.txtNumInicialBS.Text
        .nroSerie = Me.txtNroSerieNotaBS.Text
        .TipoNota = TipoBoleta
        .IdUsuarioAuditoria = ml_idUsuario
    End With
End Sub

Function ValidarDatosObligatorios() As Boolean
    Dim sMensaje As String
    Dim lnIndiceLista As Integer
    Dim lbSeleccionado As Boolean
    sMensaje = ""
    ValidarDatosObligatorios = False
    If Me.txtNroSerieNotaF.Text = "" Then
        sMensaje = sMensaje + "Ingrese el Nro de Serie de la Nota de Credito" + Chr(13)
    End If
    If Len(Me.txtNroSerieNotaF.Text) < 4 Then
        sMensaje = sMensaje + "El Nro de Serie de la Nota de Credito no tiene el formato correcto" + Chr(13)
    End If
    If Me.txtNumFinalF.Text = "" Then
        sMensaje = sMensaje + "Ingrese el Nro de documento final de la Nota de Credito" + Chr(13)
    End If
    If Me.txtNumInicialF.Text = "" Then
        sMensaje = sMensaje + "Ingrese el Nro de documento inicial de la Nota de Credito" + Chr(13)
    End If
    If Me.txtNumeroUltimoF.Text = "" Then
        sMensaje = sMensaje + "Ingrese el Nro de documento ultimo de la Nota de Credito" + Chr(13)
    End If
    If Mid(Me.txtNroSerieNotaF.Text, 1, 1) <> "F" And Mid(Me.txtNroSerieNotaF.Text, 1, 1) <> "f" Then
        sMensaje = sMensaje + "La nota de credito para facturas no tiene el formato correcto FXXX" + Chr(13)
    End If
    If Mid(Me.txtNroSerieNotaB.Text, 1, 1) <> "B" And Mid(Me.txtNroSerieNotaB.Text, 1, 1) <> "b" Then
        sMensaje = sMensaje + "La nota credito para boletas no tiene el formato correcto BXXX" + Chr(13)
    End If
      
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   
   ValidarDatosObligatorios = True
End Function


Private Sub btnGrabar_Click()
    If ValidarDatosObligatorios() Then
        CargaDatosAlObjetosDeDatos
        If ModificarDatos() Then
            MsgBox "Los datos de la nota de credito se modificaron satisfactoriamente", vbInformation, Me.Caption
            Me.Visible = False
        Else
            MsgBox "No se pudo modificar los datos de la nota de credito" + Chr(13), vbExclamation, Me.Caption
        End If
    End If
End Sub

Function ModificarDatos() As Boolean
    If mo_AdminCaja.NotaCreditoDebitoTipoNotaModificar(oDONotaCreditoDebitoTipoNotaF) Then
        If mo_AdminCaja.NotaCreditoDebitoTipoNotaModificar(oDONotaCreditoDebitoTipoNotaB) Then
            If mo_AdminCaja.NotaCreditoDebitoTipoNotaModificar(oDONotaCreditoDebitoTipoNotaFS) Then
                If mo_AdminCaja.NotaCreditoDebitoTipoNotaModificar(oDONotaCreditoDebitoTipoNotaBS) Then
                     ModificarDatos = True
                End If
            End If
        End If
    End If
End Function

Private Sub cmdEliminaRD_Click()
    Dim lbSeguir As Boolean
    lbSeguir = True
    If lbTieneTicketSunat = True Then
       If MsgBox("Ya tiene grabado el TICKET SUNAT" & Chr(13) & "¿Está seguro de eliminar?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
          lbSeguir = False
       End If
    Else
        If MsgBox("¿Está seguro?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
           lbSeguir = False
        End If
    End If
    If lbSeguir = True Then
       mo_AdminCaja.Sunat_ResumenDiarioEliminarPorFechaEmision lcFechaEmision
       Set oRsTmpDctos = mo_AdminCaja.Sunat_ResumenDiarioSeleccionarPorFechaEmision(lcFechaEmision)
       Set grdResumenDiario.DataSource = oRsTmpDctos
    End If
End Sub

Private Sub cmdGrabaTicket_Click()
    Dim lbSeguir As Boolean
    lbSeguir = True
    txtTicketSunat.Text = Trim(txtTicketSunat.Text)
    If txtTicketSunat.Text = "" Then
       MsgBox "Tiene que ingresar el TICKET SUNAT", vbInformation, ""
       lbSeguir = False
    End If
    If lbTieneTicketSunat = True Then
       If MsgBox("Ya tiene grabado el TICKET SUNAT" & Chr(13) & "¿Está seguro de cambiar el TICKET?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
          lbSeguir = False
       End If
    End If
    If lbSeguir = True Then
       Dim oDoSunatResumenDia As New DoSunatResumenDia
       Dim oSunatParteDia As New SunatParteDia
       Dim oConexion As New Connection
       sighEntidades.AbreConexionSIGH oConexion
       Set oSunatParteDia.Conexion = oConexion
       If oRsTmpDctos.RecordCount > 0 Then
            oRsTmpDctos.MoveFirst
            Do While Not oRsTmpDctos.EOF
                 oDoSunatResumenDia.ID = oRsTmpDctos!ID
                 If oSunatParteDia.SeleccionarPorId(oDoSunatResumenDia) Then
                    oDoSunatResumenDia.DctoSunat = Me.txtTicketSunat.Text
                    If oSunatParteDia.Modificar(oDoSunatResumenDia) Then
                    End If
                 End If
                 oRsTmpDctos.MoveNext
            Loop
       End If
       oConexion.Close
       Set oDoSunatResumenDia = Nothing
       Set oSunatParteDia = Nothing
       Set oConexion = Nothing
       '
       Set oRsTmpDctos = mo_AdminCaja.Sunat_ResumenDiarioSeleccionarPorFechaEmision(lcFechaEmision)
       Set grdResumenDiario.DataSource = oRsTmpDctos
       MsgBox "Se guardó TICKET correctamente", vbInformation, ""
       Me.Visible = False
    End If
End Sub

Private Sub cmdSoloMuestra_Click()
    If chkUsaResumenDiario.Value = 1 Then
        Me.MousePointer = 11
'        Dim oExportar As New SIGHProxies.Procesos
'        Dim lcTempo As Object
'        Set lcTempo = CreateObject("Scripting.FileSystemObject")
'        sighentidades.Parametro378 = "c:\tmpDebb"
'        lcTempo.CreateFolder sighentidades.Parametro378
'        oExportar.ExportarFacturasBoletas txtFechaInicio.Text, txtFechaFin.Text, "", "", True
'        lcTempo.deleteFolder sighentidades.Parametro378
'        sighentidades.Parametro378 = lcBuscaParametro.SeleccionaFilaParametro(378)
        '
        
        lbTieneTicketSunat = False
        Frame(3).Visible = False
        Me.cmdEliminaRD.Visible = False
        txtTicketSunat.Text = ""
        lcFechaEmision = Year(CDate(Me.txtFechaInicio.Text)) & "-" & Right("00" & Month(CDate(Me.txtFechaInicio.Text)), 2) & "-" & _
                         Right("00" & Day(CDate(Me.txtFechaInicio.Text)), 2)
        Set oRsTmpDctos = mo_AdminCaja.Sunat_ResumenDiarioSeleccionarPorFechaEmision(lcFechaEmision)
        If oRsTmpDctos.RecordCount > 0 Then
           Me.cmdEliminaRD.Visible = True
           Frame(3).Visible = True
           If Not IsNull(oRsTmpDctos!DctoSunat) Then
                txtTicketSunat.Text = oRsTmpDctos!DctoSunat
                lbTieneTicketSunat = True
           End If
        Else
           MsgBox "No tiene datos grabados", vbInformation, ""
        End If
        Set grdResumenDiario.DataSource = oRsTmpDctos
        mo_Apariencia.ConfigurarFilasBiColores grdResumenDiario, sighEntidades.GrillaConFilasBicolor
        lbNoPulsoBuscarHistorico = True
        Me.MousePointer = 1
        
        'Set oExportar = Nothing
    End If

End Sub

Private Sub Command1_Click()
    Me.Visible = False
End Sub

Private Sub Command2_Click()
    Me.Visible = False
End Sub

Private Sub btnGuardarConfFacturador_Click()
    Dim oConexion As New ADODB.Connection
    Dim oDOPArametro As New DOPArametro
    Dim oParametros As New Parametros
    oConexion.Open sighEntidades.CadenaConexion
    Set oParametros.Conexion = oConexion
    
    If Me.txtDataFacturador.Text = "" Then
        MsgBox "Ingrese la ruta donde se guardara los archivos planos", vbInformation, ""
        Exit Sub
    End If
    
    oDOPArametro.IdParametro = 378
    If oParametros.SeleccionarPorId(oDOPArametro) Then
        oDOPArametro.ValorTexto = Trim(Me.txtDataFacturador.Text)
        If oParametros.Modificar(oDOPArametro) Then
            MsgBox "Se guardo la ruta de la data para el facturador", vbInformation, ""
        End If
    End If
    Set oParametros = Nothing
End Sub

Private Sub Form_Load()
    ldHoy = CDate(lcBuscaParametro.RetornaFechaServidorSQL)
    Me.chkUsaResumenDiario.Value = IIf(lcBuscaParametro.SeleccionaFilaParametro(571) = "S", 1, 0)
    If Me.chkUsaResumenDiario.Value = 1 Then
       Me.cmdSoloMuestra.Visible = True
       Me.grdResumenDiario.Visible = True
       Me.Frame(2).Visible = True
    End If
    Dim lcMensajeLicencia As String, lbTieneLicenciaParaNotaCreditoYsunat As Boolean
    lbTieneLicenciaParaNotaCreditoYsunat = True
    btnAceptar.Visible = lbTieneLicenciaParaNotaCreditoYsunat
    
    
    Me.txtFechaInicio.Text = Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY) & " 00:01"
    Me.txtFechaFin.Text = Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY) & " 23:59"
    
    'FACTURA
    Dim oDONotaCreditoDebitoTipoNotaF As New DONotaCreditoDebitoTipoNota
    Set oDONotaCreditoDebitoTipoNotaF = mo_AdminCaja.NotaCreditoDebitoTipoNotaSeleccionarPorTipo(TipoNota, TipoFactura)   'Nota Credito
    Me.txtNroSerieNotaF.Text = oDONotaCreditoDebitoTipoNotaF.nroSerie
    Me.txtNumInicialF.Text = oDONotaCreditoDebitoTipoNotaF.NroDocumentoInicial
    Me.txtNumFinalF.Text = oDONotaCreditoDebitoTipoNotaF.NroDocumentoFinal
    Me.txtNumeroUltimoF.Text = oDONotaCreditoDebitoTipoNotaF.nrodocumento
    'BOLETA
    Dim oDONotaCreditoDebitoTipoNotaB As New DONotaCreditoDebitoTipoNota
    Set oDONotaCreditoDebitoTipoNotaB = mo_AdminCaja.NotaCreditoDebitoTipoNotaSeleccionarPorTipo(TipoNota, TipoBoleta)   'Nota Credito
    Me.txtNroSerieNotaB.Text = oDONotaCreditoDebitoTipoNotaB.nroSerie
    Me.txtNumInicialB.Text = oDONotaCreditoDebitoTipoNotaB.NroDocumentoInicial
    Me.txtNumFinalB.Text = oDONotaCreditoDebitoTipoNotaB.NroDocumentoFinal
    Me.txtNumeroUltimoB.Text = oDONotaCreditoDebitoTipoNotaB.nrodocumento
    'FACTURA-servicio
    Set oDONotaCreditoDebitoTipoNotaF = mo_AdminCaja.NotaCreditoDebitoTipoNotaSeleccionarPorTipo(TipoNotaServicio, TipoFactura)  'Nota Credito
    Me.txtNroSerieNotaS.Text = oDONotaCreditoDebitoTipoNotaF.nroSerie
    Me.txtNumInicialS.Text = oDONotaCreditoDebitoTipoNotaF.NroDocumentoInicial
    Me.txtNumFinalS.Text = oDONotaCreditoDebitoTipoNotaF.NroDocumentoFinal
    Me.txtNumeroUltimoS.Text = oDONotaCreditoDebitoTipoNotaF.nrodocumento
    'BOLETA-servicio
    Set oDONotaCreditoDebitoTipoNotaB = mo_AdminCaja.NotaCreditoDebitoTipoNotaSeleccionarPorTipo(TipoNotaServicio, TipoBoleta)   'Nota Credito
    Me.txtNroSerieNotaBS.Text = oDONotaCreditoDebitoTipoNotaB.nroSerie
    Me.txtNumInicialBS.Text = oDONotaCreditoDebitoTipoNotaB.NroDocumentoInicial
    Me.txtNumFinalBS.Text = oDONotaCreditoDebitoTipoNotaB.NroDocumentoFinal
    Me.txtNumeroUltimoBS.Text = oDONotaCreditoDebitoTipoNotaB.nrodocumento
    
    
    Me.txtDataFacturador.Text = lcBuscaParametro.SeleccionaFilaParametro(378)
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

Private Sub txtFechaFin_LostFocus()
If Not IsDate(txtFechaFin.Text) Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFechaFin.Text = sighEntidades.FECHA_VACIA_DMY_HM
        Exit Sub
    End If
    lbNoPulsoBuscarHistorico = False
End Sub

Private Sub txtFechaInicio_LostFocus()
If Not IsDate(txtFechaInicio.Text) Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFechaInicio.Text = sighEntidades.FECHA_VACIA_DMY_HM
        Exit Sub
    End If
End Sub

Private Sub txtNroSerieNotaB_LostFocus()
    Dim oRsNota As New Recordset
    Set oRsNota = mo_AdminCaja.NotaCreditoDebitoSeleccionarPorNroSerie(txtNroSerieNotaB.Text, TipoBoleta)
    If oRsNota.RecordCount > 0 Then
        oRsNota.MoveFirst
        If oRsNota.Fields!nroSerie = oRsNota.Fields!NroSerieTipo Then
            Me.txtNumFinalB.Text = Right("00000000" & Trim(oRsNota.Fields!NroDocFinalTipo), 8)
            Me.txtNumInicialB.Text = Right("00000000" & Trim(oRsNota.Fields!NroDocInicioTipo), 8)
            Me.txtNumeroUltimoB.Text = Right("00000000" & Trim(oRsNota.Fields!NroDocUltimoTipo), 8)
        Else
            Me.txtNumeroUltimoB.Text = Right("00000000" & oRsNota.Fields!nrodocumento, 8)
            Me.txtNumFinalB.Text = "99999999"
            Me.txtNumInicialB.Text = "00000000"
        End If
    Else
        Me.txtNumFinalB.Text = "99999999"
        Me.txtNumInicialB.Text = "00000000"
        Me.txtNumeroUltimoB.Text = "00000000"
    End If
End Sub

Private Sub txtNroSerieNotaF_LostFocus()
    Dim oRsNota As New Recordset
    Set oRsNota = mo_AdminCaja.NotaCreditoDebitoSeleccionarPorNroSerie(txtNroSerieNotaF.Text, TipoFactura)
    If oRsNota.RecordCount > 0 Then
        oRsNota.MoveFirst
        If oRsNota.Fields!nroSerie = oRsNota.Fields!NroSerieTipo Then
            Me.txtNumFinalF.Text = Right("00000000" & Trim(oRsNota.Fields!NroDocFinalTipo), 8)
            Me.txtNumInicialF.Text = Right("00000000" & Trim(oRsNota.Fields!NroDocInicioTipo), 8)
            Me.txtNumeroUltimoF.Text = Right("00000000" & Trim(oRsNota.Fields!NroDocUltimoTipo), 8)
        Else
            Me.txtNumeroUltimoF.Text = Right("00000000" & oRsNota.Fields!nrodocumento, 8)
            Me.txtNumFinalF.Text = "99999999"
            Me.txtNumInicialF.Text = "00000000"
        End If
    Else
        Me.txtNumFinalF.Text = "99999999"
        Me.txtNumInicialF.Text = "00000000"
        Me.txtNumeroUltimoF.Text = "00000000"
    End If
End Sub

Private Sub txtNumFinalF_LostFocus()
    If Len(txtNumFinalF.Text) < 8 Then
        txtNumFinalF.Text = Right("00000000" & txtNumFinalF.Text, 8)
    End If
End Sub

Private Sub txtNumInicialF_LostFocus()
    If Len(txtNumInicialF.Text) < 8 Then
        txtNumInicialF.Text = Right("00000000" & txtNumInicialF.Text, 8)
    End If
End Sub

Private Sub txtNumeroUltimoF_LostFocus()
    If Len(txtNumeroUltimoF.Text) < 8 Then
        txtNumeroUltimoF.Text = Right("00000000" & txtNumeroUltimoF.Text, 8)
    End If
End Sub

Private Sub txtNumFinalB_LostFocus()
    If Len(txtNumFinalB.Text) < 8 Then
        txtNumFinalB.Text = Right("00000000" & txtNumFinalB.Text, 8)
    End If
End Sub

Private Sub txtNumInicialB_LostFocus()
    If Len(txtNumInicialB.Text) < 8 Then
        txtNumInicialB.Text = Right("00000000" & txtNumInicialB.Text, 8)
    End If
End Sub

Private Sub txtNumeroUltimoB_LostFocus()
    If Len(txtNumeroUltimoB.Text) < 8 Then
        txtNumeroUltimoB.Text = Right("00000000" & txtNumeroUltimoB.Text, 8)
    End If
End Sub



