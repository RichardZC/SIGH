VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form CajaApruebaNotaCredito 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aprobación de Nota de Credito"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11580
   Icon            =   "CajaApruebaNotaCredito.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tabNotaCredito 
      Height          =   8655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   15266
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "APROBACIÓN NOTA DE CREDITO"
      TabPicture(0)   =   "CajaApruebaNotaCredito.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame Frame2 
         Height          =   1080
         Index           =   0
         Left            =   120
         TabIndex        =   73
         Top             =   7500
         Width           =   11295
         Begin VB.CommandButton cmdImpTicket 
            Caption         =   "Imp.Atención Ticket"
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
            Left            =   1545
            Picture         =   "CajaApruebaNotaCredito.frx":08E6
            Style           =   1  'Graphical
            TabIndex        =   77
            Top             =   240
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.CommandButton btnAceptar 
            Caption         =   "Aceptar (F2)"
            DisabledPicture =   "CajaApruebaNotaCredito.frx":0DBF
            DownPicture     =   "CajaApruebaNotaCredito.frx":121F
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
            Left            =   4320
            Picture         =   "CajaApruebaNotaCredito.frx":1694
            Style           =   1  'Graphical
            TabIndex        =   76
            Top             =   240
            Width           =   1365
         End
         Begin VB.CommandButton btnCancelar 
            Cancel          =   -1  'True
            Caption         =   "Cancelar (ESC)"
            DisabledPicture =   "CajaApruebaNotaCredito.frx":1B09
            DownPicture     =   "CajaApruebaNotaCredito.frx":1FCD
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
            Left            =   5760
            Picture         =   "CajaApruebaNotaCredito.frx":24B9
            Style           =   1  'Graphical
            TabIndex        =   75
            Top             =   225
            Width           =   1365
         End
         Begin VB.CommandButton btnImprimeNotaCredito 
            Caption         =   "Imp.Atención"
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
            Left            =   120
            Picture         =   "CajaApruebaNotaCredito.frx":29A5
            Style           =   1  'Graphical
            TabIndex        =   74
            Top             =   240
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.PictureBox ucGestionCaja1 
            Height          =   375
            Left            =   9810
            ScaleHeight     =   315
            ScaleWidth      =   810
            TabIndex        =   78
            Top             =   300
            Visible         =   0   'False
            Width           =   870
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7215
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   11295
         Begin VB.Frame Frame3 
            Caption         =   "Nota de Credito"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Width           =   11055
            Begin VB.ComboBox cmbEstadoNota 
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
               Left            =   5760
               Style           =   2  'Dropdown List
               TabIndex        =   67
               Top             =   120
               Visible         =   0   'False
               Width           =   390
            End
            Begin VB.TextBox txtNroSerie 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   405
               Left            =   120
               Locked          =   -1  'True
               MaxLength       =   4
               TabIndex        =   66
               Top             =   480
               Width           =   1215
            End
            Begin VB.TextBox txtNroDocumento 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   405
               Left            =   1320
               Locked          =   -1  'True
               MaxLength       =   12
               TabIndex        =   65
               Top             =   480
               Width           =   1815
            End
            Begin VB.Label Label8 
               Caption         =   "Nº Serie"
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
               TabIndex        =   72
               Top             =   240
               Width           =   825
            End
            Begin VB.Label Label9 
               Caption         =   "Nº Documento"
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
               Left            =   1320
               TabIndex        =   71
               Top             =   240
               Width           =   1245
            End
            Begin VB.Label Label10 
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
               Height          =   255
               Left            =   3120
               TabIndex        =   70
               Top             =   240
               Width           =   1245
            End
            Begin VB.Label lblNroOrden 
               Caption         =   "Nº Orden: "
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
               Height          =   225
               Left            =   8640
               TabIndex        =   69
               Top             =   0
               Width           =   2385
            End
            Begin VB.Label lblEstadoNota 
               Alignment       =   2  'Center
               BackColor       =   &H80000009&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   405
               Left            =   3120
               TabIndex        =   68
               Top             =   480
               Width           =   2535
            End
         End
         Begin VB.Frame Frame 
            Caption         =   "Detalle la nota de crédito"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3015
            Index           =   0
            Left            =   120
            TabIndex        =   46
            Top             =   4080
            Width           =   11055
            Begin VB.TextBox txtRazonSocial 
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
               MaxLength       =   50
               TabIndex        =   57
               Top             =   600
               Width           =   4875
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
               Left            =   5040
               MaxLength       =   11
               TabIndex        =   56
               Top             =   600
               Width           =   2235
            End
            Begin VB.TextBox txtDireccion 
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
               MaxLength       =   50
               TabIndex        =   55
               Top             =   1200
               Width           =   4875
            End
            Begin VB.ComboBox cmbMotivo 
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
               Left            =   5040
               Style           =   2  'Dropdown List
               TabIndex        =   54
               Top             =   1200
               Width           =   4575
            End
            Begin VB.Frame DetalleNotaCredito 
               BorderStyle     =   0  'None
               Height          =   1215
               Left            =   120
               TabIndex        =   49
               Top             =   1680
               Width           =   10815
               Begin VB.TextBox txtObservaciones 
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
                  Height          =   795
                  Left            =   0
                  MaxLength       =   500
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   51
                  Top             =   360
                  Width           =   8985
               End
               Begin VB.TextBox txtTotal 
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
                  ForeColor       =   &H00000000&
                  Height          =   795
                  Left            =   9000
                  MultiLine       =   -1  'True
                  TabIndex        =   50
                  Top             =   360
                  Width           =   1815
               End
               Begin VB.Label Label7 
                  Alignment       =   2  'Center
                  BackColor       =   &H8000000C&
                  Caption         =   "Concepto"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000E&
                  Height          =   375
                  Left            =   0
                  TabIndex        =   53
                  Top             =   0
                  Width           =   8985
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H8000000C&
                  Caption         =   "Importe (S./)"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000E&
                  Height          =   375
                  Left            =   9000
                  TabIndex        =   52
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin VB.OptionButton opcAnulaTotal 
               Caption         =   "Total"
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
               Left            =   9840
               TabIndex        =   48
               Top             =   840
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton opcAnulaParcial 
               Caption         =   "Parcial"
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
               Left            =   9840
               TabIndex        =   47
               Top             =   1200
               Width           =   975
            End
            Begin VB.Label Label17 
               Caption         =   "Razón Social   ó   Apellidos y Nombres"
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
               Left            =   120
               TabIndex        =   63
               Top             =   360
               Width           =   5685
            End
            Begin VB.Label Label12 
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
               Height          =   255
               Left            =   5040
               TabIndex        =   62
               Top             =   360
               Width           =   1245
            End
            Begin VB.Label Label19 
               Caption         =   "Dirección"
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
               Left            =   120
               TabIndex        =   61
               Top             =   960
               Width           =   5685
            End
            Begin VB.Label Label21 
               Caption         =   "Motivo"
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
               Left            =   5040
               TabIndex        =   60
               Top             =   960
               Width           =   1245
            End
            Begin VB.Label Label13 
               Caption         =   "Fecha de emisión"
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
               Left            =   7320
               TabIndex        =   59
               Top             =   360
               Width           =   1845
            End
            Begin VB.Label lblFechaNota 
               Alignment       =   2  'Center
               BackColor       =   &H80000009&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   315
               Left            =   7320
               TabIndex        =   58
               Top             =   600
               Width           =   2295
            End
         End
         Begin VB.Frame Frame 
            Caption         =   "Comprobante de pago afectado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2775
            Index           =   2
            Left            =   120
            TabIndex        =   2
            Top             =   1200
            Width           =   11055
            Begin VB.Frame fraNotaIngresoFarm 
               Caption         =   "Nota de Ingreso a Famarcia"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   1680
               TabIndex        =   25
               Top             =   1560
               Visible         =   0   'False
               Width           =   9255
               Begin VB.TextBox txtMovimientoFarm 
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
                  MaxLength       =   9
                  TabIndex        =   31
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.CommandButton btnLimpiarNotaCredito 
                  Height          =   315
                  Left            =   7920
                  Picture         =   "CajaApruebaNotaCredito.frx":2E7E
                  Style           =   1  'Graphical
                  TabIndex        =   30
                  Top             =   600
                  Width           =   1305
               End
               Begin VB.CommandButton btnVistaNotaCredito 
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
                  Left            =   2880
                  Picture         =   "CajaApruebaNotaCredito.frx":34A7
                  Style           =   1  'Graphical
                  TabIndex        =   29
                  Top             =   240
                  Width           =   315
               End
               Begin VB.TextBox lblFechaMov 
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
                  Left            =   5160
                  TabIndex        =   28
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.TextBox lblTotDevuelto 
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
                  Left            =   7920
                  TabIndex        =   27
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.TextBox lblDetalleNotaIngrFarm 
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
                  TabIndex        =   26
                  Top             =   600
                  Width           =   6135
               End
               Begin VB.Label Label22 
                  Caption         =   "Nº Movimiento"
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
                  TabIndex        =   35
                  Top             =   290
                  Width           =   1305
               End
               Begin VB.Label Label20 
                  Caption         =   "Farmacia Destino"
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
                  TabIndex        =   34
                  Top             =   615
                  Width           =   1515
               End
               Begin VB.Label Label24 
                  Caption         =   "Fecha de Movimiento"
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
                  Left            =   3360
                  TabIndex        =   33
                  Top             =   290
                  Width           =   1785
               End
               Begin VB.Label Label25 
                  Caption         =   "Total Devuelto"
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
                  Left            =   6600
                  TabIndex        =   32
                  Top             =   285
                  Width           =   1305
               End
            End
            Begin VB.ComboBox cmbTipoComprobante 
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
               Left            =   10320
               Style           =   2  'Dropdown List
               TabIndex        =   24
               Top             =   120
               Visible         =   0   'False
               Width           =   390
            End
            Begin VB.TextBox txtSerieComprobante 
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
               MaxLength       =   4
               TabIndex        =   22
               Top             =   600
               Width           =   1215
            End
            Begin VB.TextBox txtDocumentoComprobante 
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
               MaxLength       =   8
               TabIndex        =   23
               Top             =   600
               Width           =   1815
            End
            Begin VB.CommandButton btnLimpiarDocumento 
               Height          =   315
               Left            =   9620
               Picture         =   "CajaApruebaNotaCredito.frx":3A31
               Style           =   1  'Graphical
               TabIndex        =   21
               Top             =   1200
               Width           =   1305
            End
            Begin VB.CommandButton btnVistaBoletaServicio 
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
               Picture         =   "CajaApruebaNotaCredito.frx":405A
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   1800
               Visible         =   0   'False
               Width           =   315
            End
            Begin VB.Frame fraServicioCita 
               Caption         =   "Consulta Externa - Cita"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   1680
               TabIndex        =   11
               Top             =   1560
               Visible         =   0   'False
               Width           =   9255
               Begin VB.TextBox txtServicioCE 
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
                  Left            =   1200
                  TabIndex        =   15
                  Top             =   240
                  Width           =   4335
               End
               Begin VB.TextBox txtMedicoCE 
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
                  Left            =   1200
                  TabIndex        =   14
                  Top             =   600
                  Width           =   4335
               End
               Begin VB.TextBox txtFechaCE 
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
                  Left            =   6720
                  TabIndex        =   13
                  Top             =   240
                  Width           =   2175
               End
               Begin VB.TextBox txtTurnoCE 
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
                  Left            =   6720
                  TabIndex        =   12
                  Top             =   600
                  Width           =   2175
               End
               Begin VB.Image ImgAdvertencia 
                  Height          =   210
                  Left            =   2160
                  Picture         =   "CajaApruebaNotaCredito.frx":45E4
                  Top             =   0
                  Width           =   225
               End
               Begin VB.Label Label30 
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
                  Height          =   255
                  Left            =   5760
                  TabIndex        =   19
                  Top             =   615
                  Width           =   900
               End
               Begin VB.Label Label29 
                  Caption         =   "Fecha"
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
                  Left            =   5760
                  TabIndex        =   18
                  Top             =   285
                  Width           =   900
               End
               Begin VB.Label Label27 
                  Caption         =   "Médico"
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
                  TabIndex        =   17
                  Top             =   615
                  Width           =   900
               End
               Begin VB.Label Label23 
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
                  Height          =   255
                  Left            =   240
                  TabIndex        =   16
                  Top             =   285
                  Width           =   900
               End
            End
            Begin VB.TextBox txtTipoComprobante 
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
               Left            =   3120
               TabIndex        =   10
               Top             =   600
               Width           =   1935
            End
            Begin VB.TextBox txtFechoraComprob 
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
               Left            =   5040
               TabIndex        =   9
               Top             =   600
               Width           =   2295
            End
            Begin VB.TextBox txtNroCuenta 
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
               Left            =   7320
               TabIndex        =   8
               Top             =   600
               Width           =   2295
            End
            Begin VB.TextBox txtNroHistoria 
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
               TabIndex        =   7
               Top             =   600
               Width           =   1335
            End
            Begin VB.TextBox lblRazonSocial 
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
               TabIndex        =   6
               Top             =   1200
               Width           =   4935
            End
            Begin VB.TextBox lblRuc 
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
               Left            =   5040
               TabIndex        =   5
               Top             =   1200
               Width           =   2295
            End
            Begin VB.TextBox lblTotal 
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
               Left            =   7320
               TabIndex        =   4
               Top             =   1200
               Width           =   2295
            End
            Begin VB.TextBox txtTipoOrden 
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
               Height          =   315
               Left            =   120
               TabIndex        =   3
               Top             =   1800
               Width           =   1215
            End
            Begin VB.Label Label14 
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
               Height          =   255
               Left            =   7320
               TabIndex        =   45
               Top             =   360
               Width           =   1125
            End
            Begin VB.Label Label18 
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
               Height          =   255
               Left            =   5040
               TabIndex        =   44
               Top             =   960
               Width           =   1365
            End
            Begin VB.Label Label4 
               Caption         =   "Razón Social   ó   Apellidos y Nombres"
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
               Left            =   120
               TabIndex        =   43
               Top             =   960
               Width           =   4485
            End
            Begin VB.Label Label15 
               Caption         =   "Nº Historia"
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
               Left            =   9600
               TabIndex        =   42
               Top             =   360
               Width           =   1125
            End
            Begin VB.Label Label3 
               Caption         =   "Orden"
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
               Left            =   120
               TabIndex        =   41
               Top             =   1560
               Width           =   1245
            End
            Begin VB.Label Label2 
               Caption         =   "Total Comprobante"
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
               Left            =   7320
               TabIndex        =   40
               Top             =   960
               Width           =   1605
            End
            Begin VB.Label Label16 
               Caption         =   "Fecha y Hora"
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
               Left            =   5040
               TabIndex        =   39
               Top             =   360
               Width           =   1605
            End
            Begin VB.Label Label11 
               Caption         =   "Tipo Comprobante"
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
               Left            =   3120
               TabIndex        =   38
               Top             =   360
               Width           =   1605
            End
            Begin VB.Label Label6 
               Caption         =   "Nº Documento"
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
               Left            =   1320
               TabIndex        =   37
               Top             =   360
               Width           =   1245
            End
            Begin VB.Label Label5 
               Caption         =   "Nº Serie"
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
               TabIndex        =   36
               Top             =   360
               Width           =   825
            End
         End
      End
   End
End
Attribute VB_Name = "CajaApruebaNotaCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: MINSA - OGEI - CDI
'        Aplicativo: SisGalenPlus v.3
'        Programa: Aprueba Nota de Credito
'        Programado por: FRANKLIN CACHAY V.
'        Fecha: Julio 2015
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim ml_IdUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdProgramacion As Long
Dim ml_Estado As Integer
Dim mo_cmbEstadoNota As New SIGHEntidades.ListaDespleglable
Dim mo_cmbTipoComprobante As New SIGHEntidades.ListaDespleglable
Dim mo_cmbMotivo As New SIGHEntidades.ListaDespleglable
Dim mb_SeHaModificadoNota As Boolean
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim oRsBusquedaRecibos As New ADODB.Recordset
Dim oRsDatosCita As New ADODB.Recordset
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim ml_idTipoNota As Integer
Dim mc_TipoNota As String
Dim ml_idTipoOrden As Integer
Dim oDoNotaCreditoDebito As New DoNotaCreditoDebito
Dim ml_idRegistroSeleccionado As Long
Dim ORsDetalleComprobante As New Recordset
Const TipoFactura As String = "2"
Const TipoBoleta As String = "3"
Dim lnTotalBoleta As Double
Dim lnIdCaja As Long, lnIdGestionCaja As Long, lnIdTurno As Long
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
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
Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
End Property
Property Get IdUsuario() As Long
   IdUsuario = ml_IdUsuario
End Property
Property Let idTipoNota(lValue As Long)
   ml_idTipoNota = lValue
End Property
Property Get idTipoNota() As Long
   idTipoNota = ml_idTipoNota
End Property
Property Get SeHaModificadoNota() As Boolean
   SeHaModificadoNota = mb_SeHaModificadoNota
End Property
Property Let idRegistroSeleccionado(lValue As Long)
   ml_idRegistroSeleccionado = lValue
End Property
Property Get idRegistroSeleccionado() As Long
   idRegistroSeleccionado = ml_idRegistroSeleccionado
End Property

Private Sub cmbIdTipoProgramacion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub btnBuscarDocumento_Click()
End Sub

Sub BuscarDocumentoAfectado()
    If Trim(txtSerieComprobante.Text) = "" Then
        MsgBox "Ingrese el Nro de Serie del comprobante", vbInformation, Me.Caption
        Exit Sub
    End If
    If Trim(txtDocumentoComprobante.Text) = "" Then
        MsgBox "Ingrese el Nro de Comprobante", vbInformation, Me.Caption
        Exit Sub
    End If
    MousePointer = 11
    RealizarBusqueda
    MousePointer = 1
End Sub

Private Sub btnImprimeNotaCredito_Click()
    ImprimeNC False
End Sub

Sub ImprimeNC(lbEsTicket As Boolean)
    Dim oReporte As New RptCaja
    oReporte.ImpresionNotaCredito Trim(txtNroSerie.Text), Trim(txtNroDocumento.Text), Trim(txtRazonSocial.Text), Trim(txtRuc.Text), _
                                    Trim(txtDireccion.Text), lblFechaNota.Caption, txtTipoComprobante.Text, Trim(txtSerieComprobante.Text), _
                                        Trim(txtDocumentoComprobante.Text), txtFechoraComprob.Text, _
                                        Trim(txtObservaciones.Text), "S/." & txtTotal.Text, "S/." & txtTotal.Text, _
                                        cmbMotivo.Text, lbEsTicket
    Set oReporte = Nothing

End Sub

Private Sub btnLimpiarDocumento_Click()
    txtSerieComprobante.Text = ""
    txtDocumentoComprobante.Text = ""
    txtTipoComprobante.Text = ""
    txtFechoraComprob.Text = ""
    txtNroCuenta.Text = ""
    txtNroHistoria.Text = ""
    lblRazonSocial.Text = ""
    lblRuc.Text = ""
    lblTotal.Text = ""
    txtTipoOrden.Text = ""
    txtTipoOrden.Tag = ""
    txtMovimientoFarm.Text = ""
    lblFechaMov.Text = ""
    lblTotDevuelto.Text = ""
    lblDetalleNotaIngrFarm.Text = ""
    mo_Formulario.HabilitarDeshabilitar txtSerieComprobante, True
    mo_Formulario.HabilitarDeshabilitar txtDocumentoComprobante, True
    btnVistaNotaCredito.Enabled = False
    fraNotaIngresoFarm.Visible = False
    fraServicioCita.Visible = False
    txtServicioCE.Text = ""
    txtMedicoCE.Text = ""
    txtFechaCE.Text = ""
    txtTurnoCE.Text = ""
    txtRazonSocial.Text = ""
    txtRuc.Text = ""
    txtDireccion.Text = ""
    mo_cmbMotivo.BoundText = 1
    cmbMotivo_Change
    txtObservaciones.Text = ""
    txtTotal.Text = ""
End Sub

Private Sub btnMovimientoFarm_Click()
End Sub

Private Sub btnLimpiarNotaCredito_Click()
    txtMovimientoFarm.Text = ""
    lblFechaMov.Text = ""
    lblTotDevuelto.Text = ""
    lblDetalleNotaIngrFarm.Text = ""
    btnVistaNotaCredito.Enabled = False
    txtObservaciones.Text = ""
    txtTotal.Text = ""
'    txtTotal.Tag = ""
End Sub

'Private Sub btnVistaDocumento_Click()
'    Select Case ml_idTipoOrden
'      Case 1  'Servicios en CAJA SERVICIO
'        ImpresionDelRecibo Trim(txtSerieComprobante.Text), Trim(txtDocumentoComprobante.Text), sghServicio, sghPantalla, IIf(mo_cmbTipoComprobante.BoundText = 2, True, False)
'      Case 2  'Bienes e insumos en CAJA FARMACIA
'        ImpresionDelRecibo Trim(txtSerieComprobante.Text), Trim(txtDocumentoComprobante.Text), sghbien, sghPantalla, IIf(mo_cmbTipoComprobante.BoundText = 2, True, False)
'      Case 3  'Bienes e insumos en CAJA SERVICIO
'        ImpresionDelRecibo Trim(txtSerieComprobante.Text), Trim(txtDocumentoComprobante.Text), sghbien, sghPantalla, IIf(mo_cmbTipoComprobante.BoundText = 2, True, False)
'    End Select
'End Sub

Private Sub btnVistaBoletaServicio_Click()
    Select Case ml_idTipoOrden
      Case 1  'Servicios en CAJA SERVICIO
        ImpresionDelRecibo Trim(txtSerieComprobante.Text), Trim(txtDocumentoComprobante.Text), sghServicio, sghPantalla, IIf(mo_cmbTipoComprobante.BoundText = 2, True, False)
      Case 2  'Bienes e insumos en CAJA FARMACIA
        ImpresionDelRecibo Trim(txtSerieComprobante.Text), Trim(txtDocumentoComprobante.Text), sghbien, sghPantalla, IIf(mo_cmbTipoComprobante.BoundText = 2, True, False)
      Case 3  'Bienes e insumos en CAJA SERVICIO
        ImpresionDelRecibo Trim(txtSerieComprobante.Text), Trim(txtDocumentoComprobante.Text), sghbien, sghPantalla, IIf(mo_cmbTipoComprobante.BoundText = 2, True, False)
    End Select
End Sub

Private Sub btnVistaNotaCredito_Click()
    If lblDetalleNotaIngrFarm.Tag = "" Then Exit Sub
    Dim mo_ReglasFarmacia As New ReglasFarmacia
    Dim oDOfarmAlmacen As New DOfarmAlmacen
    Set oDOfarmAlmacen = mo_ReglasFarmacia.FarmAlmacenSeleccionarPorId(lblDetalleNotaIngrFarm.Tag)
    Dim oRptClase As New rCrystal
    oRptClase.MovTipo = "E"
    oRptClase.Documento = ""
    oRptClase.TextoDelFiltro = "NOTA DE INGRESO"
    oRptClase.Almacen = "(" & oDOfarmAlmacen.codigoSISMED & ")" & lblDetalleNotaIngrFarm.Text
    oRptClase.AlmacenO = ""
    oRptClase.HoraInicio = lblFechaMov.Text
    oRptClase.HoraFin = ""
    oRptClase.Importe = lblTotDevuelto.Tag
    oRptClase.TipoReporte = "NiNs"
'    oRptClase.Observaciones = ""
    oRptClase.IdUsuario = ml_IdUsuario
    oRptClase.Show vbModal
    Set oRptClase = Nothing
    Set mo_ReglasFarmacia = Nothing
    Set oDOfarmAlmacen = Nothing
End Sub

Sub SeleccionaMotivo()
    If mo_cmbMotivo.BoundText = 2 Or ml_idTipoOrden = sghTipoPaqueteSoloServicio Then
        opcAnulaParcial.Visible = True
        opcAnulaTotal.Visible = True
    Else
        opcAnulaParcial.Visible = False
        opcAnulaTotal.Visible = False
    End If
End Sub

Private Sub cmbMotivo_Change()
    SeleccionaMotivo
End Sub

Private Sub cmbMotivo_KeyUp(KeyCode As Integer, Shift As Integer)
    SeleccionaMotivo
End Sub

Private Sub cmdImpTicket_Click()
    ImprimeNC True
End Sub

Private Sub Form_Initialize()
    Set mo_cmbEstadoNota.MiComboBox = cmbEstadoNota
    Set mo_cmbTipoComprobante.MiComboBox = cmbTipoComprobante
    Set mo_cmbMotivo.MiComboBox = cmbMotivo
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub
'------------------------------------------------------------------------------------

Sub ValoresPorDefecto()
    If mi_Opcion = sghAgregar Then
        Dim oDONotaCreditoDebitoTipoNota As New DONotaCreditoDebitoTipoNota
        Set oDONotaCreditoDebitoTipoNota = mo_ReglasCaja.NotaCreditoDebitoTipoNotaSeleccionarPorTipo(ml_idTipoNota, IIf(mo_cmbTipoComprobante.BoundText = "", TipoBoleta, mo_cmbTipoComprobante.BoundText))
        txtNroSerie.Text = oDONotaCreditoDebitoTipoNota.nroSerie
        txtNroDocumento.Text = Format(CLng(oDONotaCreditoDebitoTipoNota.nrodocumento) + 1, "00000000")
        Set oDONotaCreditoDebitoTipoNota = Nothing
        mo_cmbEstadoNota.BoundText = 0
        lblEstadoNota.Caption = cmbEstadoNota.Text
    End If
    lblFechaNota.Caption = lcBuscaParametro.RetornaFechaServidorSQL & " " & lcBuscaParametro.RetornaHoraServidorSQLserverFormatoGalenhos
    
    mo_cmbMotivo.BoundText = 1
    cmbMotivo_Change
    opcAnulaTotal.Value = True
    mo_Formulario.HabilitarDeshabilitar txtTotal, False
    mo_cmbEstadoNota.BoundText = 0
End Sub

Sub SkinConfigura()
  On Error GoTo ErrSkin
'  Skin1.LoadSkin WxSkin
'  Skin1.ApplySkin Me.hwnd
'  btnLimpiarDocumento.Style = 0
'  btnLimpiarNotaCredito.Style = 0
'  btnImprimeNotaCredito.Style = 0
'  btnAceptar.Style = 0
'  btnCancelar.Style = 0
ErrSkin:
End Sub

Sub Form_Load()
    SkinConfigura
    
'    Dim lcMensajeLicencia As String
'    If mo_AdminServiciosComunes.EESSconDerechosAmejoras(2, "61007", lcMensajeLicencia) = False Then
'       Me.Visible = False
'    End If
    
    mc_TipoNota = IIf(ml_idTipoNota = 2, "Nota de Crédito", "Nota de Débito")
    Select Case mi_Opcion
        Case sghAgregar
            Me.Caption = "Agregar " & mc_TipoNota
            
        Case sghModificar
            Me.Caption = "Modificar " & mc_TipoNota
        Case sghConsultar
            Me.Caption = "Consultar " & mc_TipoNota
        Case sghEliminar
            Me.Caption = "Eliminar " & mc_TipoNota
    End Select
    CargarComboBoxes
    CargarDatosAlFormulario
    mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
    CargaDatosParaPagoAutomatico
End Sub

Sub Form_Activate()
'   If mi_Opcion <> sghAgregar Then
'       If Not mb_ExistenDatos Then
'           Me.Visible = False
'           LimpiarVariablesDeMemoria
'       End If
'   Else
''        If CDate(txtFechaIni.Text) < lcBuscaParametro.RetornaFechaServidorSQL Then
''          ' MsgBox "Sólo puede programar Fechas mayores a " & lcBuscaParametro.RetornaFechaServidorSQL, vbInformation, Me.Caption
''           Dim oMensaje As New SIGHNegocios.clMensaje
''           oMensaje.MostrarFormulario "Sólo puede programar Fechas mayores a " & lcBuscaParametro.RetornaFechaServidorSQL, Me.Caption
''           Set oMensaje = Nothing
''
''
''           Me.Visible = False
''           LimpiarVariablesDeMemoria
''        End If
'   End If
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyEscape
            btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

Private Sub btnAceptar_Click()
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
                CargaDatosAlObjetosDeDatos
                If AgregarDatos() Then
                    If SIGHEntidades.Parametro378valorInt = 1 Then
'                        Me.ucGestionCaja1.PagaNotaCreditoAutomaticamente oDoNotaCreditoDebito.nroSerie, _
'                                                                         oDoNotaCreditoDebito.nrodocumento, lnIdCaja, _
'                                                                         lnIdGestionCaja, lnIdTurno
'                        ImprimeNC True
                    Else
                        MsgBox "Se registro la " & mc_TipoNota, vbInformation, Me.Caption
                        btnImprimeNotaCredito_Click
                    End If
                    Me.Visible = False
                    mb_SeHaModificadoNota = True
                    LimpiarVariablesDeMemoria
                Else
                    MsgBox "No se pudo registrar la " & mc_TipoNota + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
                End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
                CargaDatosAlObjetosDeDatos
                If ModificarDatos() Then
                    MsgBox "Se modificó la " & mc_TipoNota, vbInformation, Me.Caption
                    Me.Visible = False
                    mb_SeHaModificadoNota = True
                    LimpiarVariablesDeMemoria
                Else
                    MsgBox "No se pudo modificar la " & mc_TipoNota + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
                End If
           End If
       End If
   Case sghEliminar
            Dim oMensaje2 As New SIGHNegocios.clMensaje
            oMensaje2.MostrarFormulario Chr(13) & "Esta seguro?", Me.Caption, 20, , , True
            If oMensaje2.BotonPresionado = sghAceptar Then
                If ValidarReglas() Then
                     CargaDatosAlObjetosDeDatos
                     If ModificarDatos Then 'No se elimina fisicamente solo logicamente  If EliminarDatos() Then
                         MsgBox "Se eliminó la " & mc_TipoNota, vbInformation, Me.Caption
                         Me.Visible = False
                         mb_SeHaModificadoNota = True
                         LimpiarVariablesDeMemoria
                     Else
                         MsgBox "No se pudo eliminar la " & mc_TipoNota + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
                     End If
                End If
            End If
            Set oMensaje2 = Nothing
   End Select
End Sub

Private Sub btnCancelar_Click()
    mb_SeHaModificadoNota = False
    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub

Function ValidarDatosObligatorios() As Boolean
    Dim sMensaje As String
    ValidarDatosObligatorios = False
   
    If txtSerieComprobante.Text = "" Or txtDocumentoComprobante.Text = "" Then
        sMensaje = sMensaje + "Ingrese el número de serie y documento del comprobante afectado." + Chr(13)
    Else
        If txtTipoComprobante.Text = "" Then
            sMensaje = sMensaje + "El número de serie y documento del comprobante afectado no encontro resultados" + Chr(13)
        End If
    End If
    If ml_idTipoOrden = sghTipoPaqueteSolofarmacia Then
        If txtMovimientoFarm.Text = "" Then
            sMensaje = sMensaje + "Ingrese el número de movimiento de la nota de ingreso por devolución de medicamentos." + Chr(13)
        End If
    End If
    If txtNroSerie.Text = "" Or txtNroDocumento.Text = "" Then
        sMensaje = sMensaje + "Ingrese el número de serie y documento de la " + mc_TipoNota + "." + Chr(13)
    End If
    If txtRazonSocial.Text = "" Then
        sMensaje = sMensaje + "Ingrese la Razón Social ó Apellidos y Nombres." + Chr(13)
    End If
    If txtObservaciones.Text = "" Then
        sMensaje = sMensaje + "Ingrese el concepto" + Chr(13)
    End If
    If txtTotal.Text = "" Then
        sMensaje = sMensaje + "Ingrese el importe" + Chr(13)
    End If

   If sMensaje <> "" Then
       'MsgBox sMensaje, vbInformation, Me.Caption
       Dim oMensaje As New SIGHNegocios.clMensaje
       oMensaje.MostrarFormulario sMensaje, Me.Caption
       Set oMensaje = Nothing
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function

Function ValidarReglas() As Boolean
   Dim sMensaje As String
   ValidarReglas = False
   Dim oMensaje As New SIGHNegocios.clMensaje
   
    If ml_idTipoOrden = sghTipoPaqueteSolofarmacia Then
        If CDbl(lblTotDevuelto.Tag) > CDbl(lblTotal.Tag) Then
            oMensaje.MostrarFormulario "El total de la Nota de Ingreso (Farmacia) no puede ser mayor al total del comprobante afectado.", Me.Caption
            Set oMensaje = Nothing
            Exit Function
        End If
        If CDbl(txtTotal.Text) > CDbl(lblTotDevuelto.Tag) Then
            oMensaje.MostrarFormulario "El total de la Nota de Crédito no puede ser mayor al total de la Nota de Ingreso.", Me.Caption
            Set oMensaje = Nothing
            Exit Function
        End If
    Else
        If CDbl(txtTotal.Text) > CDbl(lblTotal.Tag) Then
            oMensaje.MostrarFormulario "El total de la Nota de Crédito no puede ser mayor al total de la Nota de Ingreso.", Me.Caption
            Set oMensaje = Nothing
            Exit Function
        End If
    End If
    
    If mi_Opcion = sghEliminar Then
        If mo_cmbEstadoNota.BoundText = 3 Then
            oMensaje.MostrarFormulario "No puede eliminar la Nota de Crédito porque ya fue canjeado en caja.", Me.Caption
            Set oMensaje = Nothing
            Exit Function
        End If
    End If

    If mi_Opcion = sghModificar Then
        If mo_cmbEstadoNota.BoundText = 3 Then
            oMensaje.MostrarFormulario "No puede modificar la Nota de Crédito porque ya fue canjeado en caja.", Me.Caption
            Set oMensaje = Nothing
            Exit Function
        End If
    End If
   
    ValidarReglas = True
End Function

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
    AgregarDatos = False
    AgregarDatos = mo_ReglasCaja.NotaCreditoDebitoAgregar(oDoNotaCreditoDebito, ml_idTipoOrden, mo_cmbTipoComprobante.BoundText)
    ms_MensajeError = mo_ReglasCaja.MensajeError
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean
    ModificarDatos = False
    ModificarDatos = mo_ReglasCaja.NotaCreditoDebitoModificar(oDoNotaCreditoDebito)
    ms_MensajeError = mo_ReglasCaja.MensajeError
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean
    EliminarDatos = False
    EliminarDatos = mo_ReglasCaja.NotaCreditoDebitoEliminar(oDoNotaCreditoDebito)
    ms_MensajeError = mo_ReglasCaja.MensajeError
End Function

Sub CargarDatosAlosControles()
    Dim oDoNotaCreditoDebito As New DoNotaCreditoDebito
    Dim orsTemp As New Recordset
    Set oDoNotaCreditoDebito = mo_AdminCaja.NotaCreditoDebitoSeleccionarPorId(ml_idRegistroSeleccionado)
    lblNroOrden.Caption = "Nº Orden: " & CStr(ml_idRegistroSeleccionado)
    txtNroSerie.Text = oDoNotaCreditoDebito.nroSerie
    txtNroDocumento.Text = oDoNotaCreditoDebito.nrodocumento
    mo_cmbEstadoNota.BoundText = oDoNotaCreditoDebito.IdEstadoNota
    lblEstadoNota.Caption = cmbEstadoNota.Text
    lblFechaNota.Caption = oDoNotaCreditoDebito.FechaAprueba
    mo_cmbMotivo.BoundText = oDoNotaCreditoDebito.idMotivo
    Set orsTemp = mo_AdminCaja.CajaComprobantesSeleccionarPorId(oDoNotaCreditoDebito.IdComprobantePago)
    If orsTemp.RecordCount > 0 Then
        txtSerieComprobante.Text = orsTemp.Fields!nroSerie
        txtDocumentoComprobante.Text = orsTemp.Fields!nrodocumento
        RealizarBusqueda
    End If
    btnLimpiarDocumento.Enabled = False
    btnLimpiarNotaCredito.Enabled = False
    txtRazonSocial.Text = oDoNotaCreditoDebito.RazonSocial
    txtRuc.Text = oDoNotaCreditoDebito.ruc
    txtDireccion.Text = oDoNotaCreditoDebito.Direccion
    txtObservaciones.Text = oDoNotaCreditoDebito.Observaciones
    txtTotal.Text = oDoNotaCreditoDebito.Total
'    txtTotal.Text = "S/. " & oDoNotaCreditoDebito.Total
    mo_Formulario.HabilitarDeshabilitar txtSerieComprobante, False
    mo_Formulario.HabilitarDeshabilitar txtDocumentoComprobante, False
    If oDoNotaCreditoDebito.TipoAnulacion = True Then
        opcAnulaTotal.Value = True
    Else
        opcAnulaParcial.Value = True
    End If
End Sub

Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_Teclado = Nothing
    Set mo_Formulario = Nothing
    Set mo_AdminServiciosComunes = Nothing
    Set mo_AdminServiciosHosp = Nothing
    Set lcBuscaParametro = Nothing
End Sub

Sub CargarComboBoxes()
    'Buscamos el comprobante
    mo_cmbEstadoNota.BoundColumn = "IdEstado"
    mo_cmbEstadoNota.ListField = "EstadoNota"
    Set mo_cmbEstadoNota.RowSource = mo_AdminCaja.NotaCreditoDebitoCargarEstadoNotaCredito

    'Codigo de Tipos de comprobante
    mo_cmbTipoComprobante.BoundColumn = "IdTipoComprobante"
    mo_cmbTipoComprobante.ListField = "Descripcion"
    Set mo_cmbTipoComprobante.RowSource = mo_AdminCaja.TiposComprobanteSeleccionarTodos
    
    mo_cmbMotivo.BoundColumn = "IdMotivo"
    mo_cmbMotivo.ListField = "Motivo"
    Set mo_cmbMotivo.RowSource = mo_AdminCaja.NotaCreditoDebitoCargarMotivo

End Sub

Sub Limpiar()
    lnTotalBoleta = 0: opcAnulaParcial.Enabled = False: Me.opcAnulaTotal.Value = True 'kike 2017
    txtTipoOrden.Tag = ""
    txtTipoOrden.Text = ""
    ml_idTipoOrden = 0
    lblTotal.Text = ""
    lblRazonSocial.Text = ""
    lblRuc.Text = ""
    txtNroCuenta.Text = ""
    txtNroHistoria.Text = ""
    txtFechoraComprob.Text = ""
    
    fraNotaIngresoFarm.Visible = False
    txtMovimientoFarm.Text = ""
    lblDetalleNotaIngrFarm.Text = ""
    ImgAdvertencia.Visible = False
    
    txtRazonSocial.Text = ""
    txtRuc.Text = ""
    txtDireccion.Text = ""
    txtObservaciones.Text = ""
    txtTotal.Text = ""
    btnVistaBoletaServicio.Visible = False
End Sub

Function ValidaComprobanteEsValidoParaAplicarNota(ByVal oRsComprobante As Recordset) As Boolean
    Dim orsTemp As New Recordset
    ValidaComprobanteEsValidoParaAplicarNota = False
    Set orsTemp = mo_AdminCaja.NotaCreditoBuscaPorIdComprobante(oRsComprobante.Fields!IdComprobantePago)
    If orsTemp.RecordCount > 0 Then
        If oRsComprobante!Total = orsTemp!Total Then
            MsgBox "El comprobante ya fue afectado por la nota de crédito " & orsTemp.Fields!nroSerie & "-" & orsTemp.Fields!nrodocumento, vbInformation, Me.Caption
            ValidaComprobanteEsValidoParaAplicarNota = True
        End If
    End If
    If oRsComprobante.Fields!IdEstadoComprobante = 9 Then
       MsgBox "El comprobante " & Trim(txtSerieComprobante.Text) & " - " & Trim(txtDocumentoComprobante.Text) & " YA ESTA ANULADO", vbInformation, Me.Caption
       ValidaComprobanteEsValidoParaAplicarNota = True
    End If
    If oRsComprobante.Fields!IdEstadoComprobante = 6 Then
        MsgBox "El comprobante " & Trim(txtSerieComprobante.Text) & " - " & Trim(txtDocumentoComprobante.Text) & " YA HA SIDO DEVUELTO", vbInformation, Me.Caption
        ValidaComprobanteEsValidoParaAplicarNota = True
    End If
    If oRsComprobante.Fields!IdEstadoComprobante = 1 Then
       MsgBox "La orden aun no ha sido PAGADA, solo se puede generar notas de crédito de ordenes PAGADAS.", vbInformation, Me.Caption
       ValidaComprobanteEsValidoParaAplicarNota = True
    End If
    
    orsTemp.Close
    Set orsTemp = Nothing
End Function

Public Sub RealizarBusqueda()
    Dim orsTemp As New Recordset
    Dim orsTemp2 As New Recordset
    Dim orsTemp3 As New Recordset
    
    Dim lcDocumentoNumero As String
    Limpiar
    'Buscamos el comprobante
    Set oRsBusquedaRecibos = Nothing
    Set oRsBusquedaRecibos = mo_AdminCaja.CajaComprobantePagoSeleccionarPorFechaOdocumento(Trim(txtSerieComprobante.Text), Trim(txtDocumentoComprobante.Text), Now, Now)
    
    If oRsBusquedaRecibos.RecordCount = 0 Then
       MsgBox "El Comprobante " & Trim(txtSerieComprobante.Text) & " - " & Trim(txtDocumentoComprobante.Text) & " NO EXISTE", vbInformation, Me.Caption
       Exit Sub
    End If
    
    If oRsBusquedaRecibos.RecordCount > 0 Then
        Do While Not oRsBusquedaRecibos.EOF
            ml_idTipoOrden = oRsBusquedaRecibos.Fields!IdTipoOrden 'Farmacia o Servicio
            'Validar que el comprobante no tenga notas de credito.
            If mi_Opcion = sghAgregar Then
                If ValidaComprobanteEsValidoParaAplicarNota(oRsBusquedaRecibos) = True Then
                    Set oRsBusquedaRecibos = Nothing
                    Exit Sub
                End If
            End If
            If ml_idTipoOrden = sghTipoPaqueteSoloServicio Then
                Set orsTemp = mo_AdminCaja.NotaCreditoConsultarOrdenServicio(oRsBusquedaRecibos.Fields!IdComprobantePago)
                If orsTemp.RecordCount > 0 Then
                    If orsTemp.Fields!idPuntoCarga = 6 Then 'CE
                        Set orsTemp2 = mo_AdminCaja.NotaCreditoConsultarCitaPorNCuenta(oRsBusquedaRecibos.Fields!idCuentaAtencion)
                        If orsTemp2.RecordCount > 0 Then
                            If mi_Opcion = sghAgregar Then
                                If Format(orsTemp2.Fields!fecha, SIGHEntidades.DevuelveFechaSoloFormato_DMY) < Format(lcBuscaParametro.RetornaFechaServidorSQL(), SIGHEntidades.DevuelveFechaSoloFormato_DMY) Or _
                                    (Format(orsTemp2.Fields!fecha, SIGHEntidades.DevuelveFechaSoloFormato_DMY) = Format(lcBuscaParametro.RetornaFechaServidorSQL(), SIGHEntidades.DevuelveFechaSoloFormato_DMY) And _
                                        orsTemp2.Fields!HoraFin < lcBuscaParametro.RetornaHoraServidorSQLserverFormatoGalenhos) Then
                                            Set orsTemp3 = mo_AdminCaja.BuscarTriaje(orsTemp2.Fields!idAtencion)
                                            If orsTemp3.RecordCount = 1 Then
                                                MsgBox "No puede modificar el comprobante " & orsTemp.Fields!nroSerie & "-" & orsTemp.Fields!nrodocumento & ".La cita ya fue atendida el " & Format(orsTemp2.Fields!fecha, SIGHEntidades.DevuelveFechaSoloFormato_DMY) & " en el servicio de " & Trim(orsTemp2.Fields!Servicio), vbInformation, Me.Caption
                                                Limpiar
                                                Set oRsBusquedaRecibos = Nothing
                                                Set orsTemp = Nothing
                                                Set orsTemp2 = Nothing
                                                Set orsTemp3 = Nothing
                                                Exit Sub
                                            End If
                                            MsgBox "Observación: La fecha " + CStr(orsTemp2.Fields!fecha) + " y hora (" + orsTemp2.Fields!HoraInicio + "-" + orsTemp2.Fields!HoraFin + ") de la Cita esta vencida, pero no se encontró atención médica." + vbCrLf + "Sugerencia: Si no tiene implementado el módulo de 'Registro de Atenciones', asegúrese que la cita no fue atendida.", vbExclamation, Me.Caption
                                            ImgAdvertencia.ToolTipText = "Observación: La fecha " + CStr(orsTemp2.Fields!fecha) + " y hora (" + orsTemp2.Fields!HoraInicio + "-" + orsTemp2.Fields!HoraFin + ") de la Cita esta vencida, pero no se encontró atención médica. Sugerencia: Si no tiene implementado el módulo de 'Registro de Atenciones', asegúrese que la cita no fue atendida."
                                            ImgAdvertencia.Visible = True
                                End If
                            End If
                            fraServicioCita.Visible = True
                            txtServicioCE.Text = orsTemp2.Fields!Servicio
                            txtMedicoCE.Text = Trim(orsTemp2.Fields!Nombres) + " " + Trim(orsTemp2.Fields!ApellidoPaterno) + " " + Trim(orsTemp2.Fields!ApellidoMaterno)
                            txtFechaCE.Text = orsTemp2.Fields!fecha
                            txtTurnoCE.Text = orsTemp2.Fields!HoraInicio + " - " + orsTemp2.Fields!HoraFin
                        Else
                            If mi_Opcion = sghAgregar Then
                                If BoletaTieneRegistradoLaboratorioImagenes(oRsBusquedaRecibos.Fields!IdComprobantePago) = True Then
                                    Set oRsBusquedaRecibos = Nothing
                                    Set orsTemp = Nothing
                                    Set orsTemp2 = Nothing
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If
                Set orsTemp = Nothing
                Set orsTemp2 = Nothing
            End If
            
            lblRazonSocial.Text = Trim(oRsBusquedaRecibos.Fields!RazonSocial)
            lblRuc.Text = IIf(IsNull(oRsBusquedaRecibos.Fields!ruc), "", oRsBusquedaRecibos.Fields!ruc)
            txtNroCuenta.Text = IIf(IsNull(oRsBusquedaRecibos.Fields!idCuentaAtencion), "", oRsBusquedaRecibos.Fields!idCuentaAtencion)
            
            txtNroHistoria.Text = ""
            If Not IsNull(oRsBusquedaRecibos.Fields!NroHistoriaClinica) Then
               txtNroHistoria.Text = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(oRsBusquedaRecibos!NroHistoriaClinica)), False)
            End If
            
            mo_cmbTipoComprobante.BoundText = oRsBusquedaRecibos.Fields!IdTipoComprobante: txtTipoComprobante.Text = cmbTipoComprobante.Text
            
            If mi_Opcion = sghAgregar Then
                Dim oDONotaCreditoDebitoTipoNota As New DONotaCreditoDebitoTipoNota
                Set oDONotaCreditoDebitoTipoNota = mo_ReglasCaja.NotaCreditoDebitoTipoNotaSeleccionarPorTipo(ml_idTipoNota, mo_cmbTipoComprobante.BoundText)
                If oDONotaCreditoDebitoTipoNota.nroSerie = "" Then
                   MsgBox "Debe configurar NOTA DE CREDITO para FACTURAS/BOlETAS en opción 'HERRAMIENTAS -> ", vbInformation, ""
                   Me.Visible = False
                   Exit Sub
                End If
                txtNroSerie.Text = oDONotaCreditoDebitoTipoNota.nroSerie
                txtNroDocumento.Text = Format(CLng(oDONotaCreditoDebitoTipoNota.nrodocumento) + 1, "00000000")
                Set oDONotaCreditoDebitoTipoNota = Nothing
                mo_cmbEstadoNota.BoundText = 0
                lblEstadoNota.Caption = cmbEstadoNota.Text
            End If
            If IsNull(oRsBusquedaRecibos.Fields!FechaCobranza) Then
               MsgBox "No se ha cobrado aún, solo se ha emitido", vbInformation, ""
               Exit Sub
            End If
            txtFechoraComprob.Text = oRsBusquedaRecibos.Fields!FechaCobranza
            txtTipoOrden.Tag = ml_idTipoOrden
            mo_Formulario.HabilitarDeshabilitar txtSerieComprobante, False
            mo_Formulario.HabilitarDeshabilitar txtDocumentoComprobante, False
            'Cuando la boleta o Factura es por Servicio, el motivo solo es Anulacion, que puede ser Total o Parcial
            If ml_idTipoOrden = sghTipoPaqueteSoloServicio Then
                txtTipoOrden.Text = "SERVICIO"
                If mi_Opcion = sghAgregar Then           'kike 2017
                    ml_idTipoNota = 3
                    ValoresPorDefecto
                
                    mo_cmbMotivo.BoundText = "1" '"2"
                    cmbMotivo_Change
                End If
                btnVistaBoletaServicio.Visible = True
                opcAnulaParcial.Enabled = True      'kike 2017
                Me.opcAnulaTotal.Value = True       'kike 2017
            End If
            'Cuando la boleta o Factura es por Farmacia, el motivo es devolucion, en este sentido tenemos que verificar la nota de ingreso a farmacia
            If ml_idTipoOrden = sghTipoPaqueteSolofarmacia Then
                txtTipoOrden.Text = "FARMACIA"
                btnVistaNotaCredito.Enabled = False
                If mi_Opcion = sghAgregar Then        'kike 2017
                    mo_cmbMotivo.BoundText = "1"
                    cmbMotivo_Change
                End If
                fraNotaIngresoFarm.Visible = True
                txtMovimientoFarm.Text = ""
                lblFechaMov.Text = ""
                lblTotDevuelto.Text = ""
                lblDetalleNotaIngrFarm.Text = ""
                lcDocumentoNumero = Trim(txtSerieComprobante.Text) + "-" + Trim(txtDocumentoComprobante.Text)
                'Consultando si hubo devolucion en farmacia
                Set orsTemp = mo_AdminCaja.NotaCreditoFarmNotaIngreso(lcDocumentoNumero)
                If orsTemp.RecordCount > 0 Then
                   Do While Not orsTemp.EOF
                       txtMovimientoFarm.Text = orsTemp.Fields!movNumero
                       mo_Formulario.HabilitarDeshabilitar txtMovimientoFarm, False
                       btnVistaNotaCredito.Enabled = True
                       lblFechaMov.Text = orsTemp.Fields!fechacreacion
                       lblTotDevuelto.Text = "S/. " & SIGHEntidades.DevuelveNumeroRedondeado(Val(orsTemp.Fields!Total))
                       lblTotDevuelto.Tag = SIGHEntidades.DevuelveNumeroRedondeado(Val(orsTemp.Fields!Total))
                       lblDetalleNotaIngrFarm.Text = orsTemp.Fields!Descripcion
                       lblDetalleNotaIngrFarm.Tag = orsTemp.Fields!idAlmacenDestino
                       orsTemp.MoveNext
                   Loop
                Else
                   MsgBox "No se ha podido encontrar la nota de ingreso a la farmacia. " + vbCrLf + "Previamente debe acercarse y hacer la devolución de medicamentos a la farmacia", vbInformation, Me.Caption
                   mo_Formulario.HabilitarDeshabilitar txtMovimientoFarm, True
                End If
            End If
            lblTotal.Text = "S/. " & oRsBusquedaRecibos.Fields!Total
            lblTotal.Tag = oRsBusquedaRecibos.Fields!Total
            lnTotalBoleta = oRsBusquedaRecibos.Fields!Total
            'Cargar datos por defecto para carga inicial de la nota de credito
            If mi_Opcion = sghAgregar Then
                txtRazonSocial.Text = Trim(oRsBusquedaRecibos.Fields!RazonSocial)
                txtRuc.Text = IIf(IsNull(oRsBusquedaRecibos.Fields!ruc), "", oRsBusquedaRecibos.Fields!ruc)
                If ml_idTipoOrden = sghTipoPaqueteSoloServicio Then
                    txtTotal.Text = lblTotal.Tag
                End If
                If ml_idTipoOrden = sghTipoPaqueteSolofarmacia Then
                    txtTotal.Text = lblTotal.Tag
                    If lblTotDevuelto.Tag <> "" Then
                        If Round(lblTotal.Tag, 0) = Round(lblTotDevuelto.Tag, 0) Then
                           txtTotal.Text = lblTotal.Tag
                        Else
                           txtTotal.Text = lblTotDevuelto.Tag
                        End If
                    End If
                End If
                txtObservaciones.Text = DevolverConceptoPorDefecto
            End If
            oRsBusquedaRecibos.MoveNext
        Loop
    End If
    Set mo_AdminCaja = Nothing
End Sub

'Sub CargarDetalleComprobanteAfectado(lnBienFarmacia As sghTipoProducto)
'    Dim rsReporte As New Recordset
'    If lnBienFarmacia = sighentidades.sghbien Then 'BIEN O MEDICAMENTI
'       Set rsReporte = mo_AdminCaja.CajaComprobantePagoProductosPorNroSerieNroDocumento(Trim(txtSerieComprobante.Text), Trim(txtDocumentoComprobante.Text))
'    Else
'       Set rsReporte = mo_AdminCaja.CajaComprobantePagoServiciosPorNroSerieNroDocumento(Trim(txtSerieComprobante.Text), Trim(txtDocumentoComprobante.Text))
'    End If
'    If oRsDetalleComprobante.State = 1 Then Set oRsDetalleComprobante = Nothing
'    With oRsDetalleComprobante
'          .Fields.Append "NombreProducto", adVarChar, 255, adFldIsNullable + adFldUpdatable
'          .Fields.Append "Cantidad", adInteger, 0, adFldIsNullable + adFldUpdatable
'          .Fields.Append "PrecioUnitario", adDouble, 0, adFldIsNullable + adFldUpdatable
'          .Fields.Append "TotalPorPagar", adDouble, 0, adFldIsNullable + adFldUpdatable
'          .CursorType = adOpenDynamic
'          .LockType = adLockOptimistic
'          .Open
'    End With
'    rsReporte.MoveFirst
'    Do While Not rsReporte.EOF
'        oRsDetalleComprobante.AddNew
'        oRsDetalleComprobante.Fields!NombreProducto = rsReporte.Fields!NombreProducto
'        oRsDetalleComprobante.Fields!Cantidad = rsReporte.Fields!Cantidad
'        oRsDetalleComprobante.Fields!PrecioUnitario = rsReporte.Fields!PrecioUnitario
'        oRsDetalleComprobante.Fields!TotalPorPagar = rsReporte.Fields!TotalPorPagar
'        oRsDetalleComprobante.Update
'        rsReporte.MoveNext
'    Loop
'End Sub

Sub CargaDatosAlObjetosDeDatos()
    With oDoNotaCreditoDebito
        .IdNota = ml_idRegistroSeleccionado
        .idTipoNota = ml_idTipoNota
        'If mi_Opcion = sghAgregar Then ValoresPorDefecto 'Carga el numero de documento final     'kike 2017
        .nroSerie = txtNroSerie.Text
        .nrodocumento = txtNroDocumento.Text
        .RazonSocial = Trim(txtRazonSocial.Text)
        .ruc = Trim(Me.txtRuc.Text)
        .Total = CCur(txtTotal.Text)
        .IdUsuarioAutoriza = ml_IdUsuario
        .FechaAprueba = Me.lblFechaNota.Caption
        .Observaciones = Me.txtObservaciones.Text
        If mi_Opcion = sghAgregar Then
            mo_cmbEstadoNota.BoundText = 1
        ElseIf mi_Opcion = sghEliminar Then
            mo_cmbEstadoNota.BoundText = 2
        End If
        .IdEstadoNota = mo_cmbEstadoNota.BoundText
        .idMotivo = mo_cmbMotivo.BoundText
        .Direccion = txtDireccion.Text
        If oRsBusquedaRecibos.RecordCount > 0 Then
            oRsBusquedaRecibos.MoveFirst
            Do While Not oRsBusquedaRecibos.EOF
                .IdComprobantePago = oRsBusquedaRecibos.Fields!IdComprobantePago
                .idPaciente = IIf(IsNull(oRsBusquedaRecibos.Fields!idPaciente), 0, oRsBusquedaRecibos.Fields!idPaciente)
                .idFarmacia = IIf(IsNull(oRsBusquedaRecibos.Fields!idFarmacia), 0, oRsBusquedaRecibos.Fields!idFarmacia)
                oRsBusquedaRecibos.MoveNext
            Loop
        End If
        .TipoAnulacion = IIf(Me.opcAnulaTotal.Value = True, True, False)
    End With
End Sub

Sub ImpresionDelRecibo(lcNroSerie As String, lcNroDcto As String, lnBienFarmacia As sghTipoProducto, lbImpresionFisica As sghImpresion, lbEsFactura As Boolean)
    Dim oRecibo As New RecibosBoleta
    oRecibo.EsAnulado = ml_Estado
    oRecibo.lbTienePermisoReimprimeBoleta = True
    oRecibo.ImprimirDEBB lcNroSerie, lcNroDcto, lnBienFarmacia, ml_IdUsuario
    oRecibo.Show 1
    Set oRecibo = Nothing
End Sub

Sub CargarDatosAlFormulario()
    btnVistaNotaCredito.Enabled = False
    mo_Formulario.HabilitarDeshabilitar txtTipoComprobante, False
    mo_Formulario.HabilitarDeshabilitar txtFechoraComprob, False
    mo_Formulario.HabilitarDeshabilitar txtNroCuenta, False
    mo_Formulario.HabilitarDeshabilitar txtNroHistoria, False
    mo_Formulario.HabilitarDeshabilitar lblRazonSocial, False
    mo_Formulario.HabilitarDeshabilitar lblRuc, False
    mo_Formulario.HabilitarDeshabilitar lblTotal, False
    
    mo_Formulario.HabilitarDeshabilitar txtTipoOrden, False
    mo_Formulario.HabilitarDeshabilitar txtServicioCE, False
    mo_Formulario.HabilitarDeshabilitar txtFechaCE, False
    mo_Formulario.HabilitarDeshabilitar txtMedicoCE, False
    mo_Formulario.HabilitarDeshabilitar txtTurnoCE, False
    
    mo_Formulario.HabilitarDeshabilitar lblFechaMov, False
    mo_Formulario.HabilitarDeshabilitar lblTotDevuelto, False
    mo_Formulario.HabilitarDeshabilitar lblDetalleNotaIngrFarm, False
    
    mo_Formulario.HabilitarDeshabilitar cmbEstadoNota, False
    mo_Formulario.HabilitarDeshabilitar cmbTipoComprobante, False
    'mo_Formulario.HabilitarDeshabilitar cmbMotivo, False
    mo_Formulario.HabilitarDeshabilitar txtTotal, False
    Select Case mi_Opcion
    Case sghAgregar
        ValoresPorDefecto
    Case sghModificar
        CargarDatosAlosControles
        btnImprimeNotaCredito.Visible = True
        Me.cmdImpTicket.Visible = True
    Case sghConsultar
        CargarDatosAlosControles
        btnAceptar.Enabled = False
        btnImprimeNotaCredito.Visible = True
        Me.cmdImpTicket.Visible = True
    Case sghEliminar
        CargarDatosAlosControles
    End Select
End Sub

Private Sub opcAnulaParcial_Click()
    If opcAnulaParcial.Value = True Then
        mo_Formulario.HabilitarDeshabilitar txtTotal, True
        If mi_Opcion = sghAgregar Then txtObservaciones.Text = DevolverConceptoPorDefecto
    End If
End Sub

Private Sub opcAnulaTotal_Click()
    If opcAnulaTotal.Value = True Then
        txtTotal.Text = lblTotal.Tag
        mo_Formulario.HabilitarDeshabilitar txtTotal, False
        If mi_Opcion = sghAgregar Then txtObservaciones.Text = DevolverConceptoPorDefecto
    End If
End Sub

Private Sub txtDocumentoComprobante_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDocumentoComprobante
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtDocumentoComprobante_KeyPress(KeyAscii As Integer)
   If Len(txtDocumentoComprobante.Text) > 0 And KeyAscii = 13 Then
      BuscarDocumentoAfectado
   ElseIf Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtMovimientoFarm_KeyPress(KeyAscii As Integer)
   If Len(txtSerieComprobante.Text) > 0 And KeyAscii = 13 Then
        Dim orsTemp As Recordset
        Dim lcDocumentoNumero As String
        lcDocumentoNumero = Trim(txtSerieComprobante.Text) + "-" + Trim(txtDocumentoComprobante.Text)
        'Consultando si hubo devolucion en farmacia
        Set orsTemp = mo_AdminCaja.NotaCreditoFarmNotaIngreso(lcDocumentoNumero)
        If orsTemp.RecordCount > 0 Then
            Do While Not orsTemp.EOF
                txtMovimientoFarm.Text = orsTemp.Fields!movNumero
                lblFechaMov.Text = orsTemp.Fields!fechacreacion
                lblTotDevuelto.Text = "S/. " & SIGHEntidades.DevuelveNumeroRedondeado(Val(orsTemp.Fields!Total))
                lblTotDevuelto.Tag = SIGHEntidades.DevuelveNumeroRedondeado(Val(orsTemp.Fields!Total))
                lblDetalleNotaIngrFarm.Text = orsTemp.Fields!Descripcion
                txtObservaciones.Text = DevolverConceptoPorDefecto
'                txtTotal.Text = lblTotDevuelto.Caption
                txtTotal.Text = lblTotDevuelto.Tag
                orsTemp.MoveNext
            Loop
        Else
            MsgBox "No se encontro nota de ingreso de farmacia con el número de movimiento registrado", vbInformation, Me.Caption
            mo_Formulario.HabilitarDeshabilitar txtMovimientoFarm, True
        End If
   ElseIf Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
        If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
            KeyAscii = 0
        End If
   End If
End Sub

Private Sub txtSerieComprobante_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtSerieComprobante
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtSerieComprobante_KeyPress(KeyAscii As Integer)
   If Len(txtSerieComprobante.Text) > 0 And txtDocumentoComprobante.Text <> "" And KeyAscii = 13 Then
      BuscarDocumentoAfectado
'   ElseIf Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
'       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
'           KeyAscii = 0
'       End If
   End If
End Sub

Function DevolverConceptoPorDefecto() As String
    DevolverConceptoPorDefecto = ""
    If ml_idTipoOrden = sghTipoPaqueteSoloServicio Then
        If fraServicioCita.Visible = True Then
            DevolverConceptoPorDefecto = "Por Devolución " + IIf(opcAnulaTotal.Value = True, opcAnulaTotal.Caption, opcAnulaParcial.Caption) + " del comprobante, "
            DevolverConceptoPorDefecto = DevolverConceptoPorDefecto + "afectando a la " + Trim(txtTipoComprobante.Text) + " " + Trim(txtSerieComprobante.Text) + "-" + Trim(txtDocumentoComprobante.Text) + "."
            DevolverConceptoPorDefecto = DevolverConceptoPorDefecto + vbCrLf + "La cita ligada al comprobante de pago se anuló "
            DevolverConceptoPorDefecto = DevolverConceptoPorDefecto + vbCrLf + "( Servicio: " + Trim(txtServicioCE.Text) + " | Turno:" & txtTurnoCE.Text
            DevolverConceptoPorDefecto = DevolverConceptoPorDefecto + " | Médico: " + Trim(txtMedicoCE.Text) + ")"
        Else
            DevolverConceptoPorDefecto = "Por Devolución " + IIf(opcAnulaTotal.Value = True, opcAnulaTotal.Caption, opcAnulaParcial.Caption) + " del comprobante, "
            DevolverConceptoPorDefecto = DevolverConceptoPorDefecto + vbCrLf + "Afectando a la " + Trim(txtTipoComprobante.Text) + " " + Trim(txtSerieComprobante.Text) + "-" + Trim(txtDocumentoComprobante.Text) + "."
        End If
    End If
    If ml_idTipoOrden = sghTipoPaqueteSolofarmacia Then
        DevolverConceptoPorDefecto = "Por devolución de medicamentos y/o insumos a la farmacia '" + lblDetalleNotaIngrFarm.Text + "'"
        DevolverConceptoPorDefecto = DevolverConceptoPorDefecto + vbCrLf + "Afectando a la " + Trim(txtTipoComprobante.Text) + " " + Trim(txtSerieComprobante.Text) + "-" + Trim(txtDocumentoComprobante.Text) + "."
        If txtMovimientoFarm.Text <> "" Then
           DevolverConceptoPorDefecto = DevolverConceptoPorDefecto + vbCrLf + "(NI por Devolución N°: " & txtMovimientoFarm.Text & ")"
        End If
    End If
End Function

Function BoletaTieneRegistradoLaboratorioImagenes(lnIdComprobantePago As Long) As Boolean
    Dim oRsTmp As New Recordset, lcSql As String
    BoletaTieneRegistradoLaboratorioImagenes = False
    Set oRsTmp = mo_AdminCaja.CajaComprobantesPagoXimagenes(lnIdComprobantePago)
    If oRsTmp.RecordCount > 0 And IsNull(oRsTmp.Fields!idCuentaAtencion) Then
       MsgBox "El Documento " & Trim(txtSerieComprobante.Text) & " - " & Trim(txtDocumentoComprobante.Text) & " tiene registrado Movimiento en Imágenes" & Chr(13) & Chr(13) & "Fecha: " & oRsTmp.Fields!fecha & ",      N° Movimiento: " & oRsTmp.Fields!IdMovimiento, vbInformation, "Caja"
       BoletaTieneRegistradoLaboratorioImagenes = True
    Else
       oRsTmp.Close
       Set oRsTmp = mo_AdminCaja.CajaComprobantesPagoXlaboratorio(lnIdComprobantePago)
       If oRsTmp.RecordCount > 0 And IsNull(oRsTmp.Fields!idCuentaAtencion) Then
           MsgBox "El Documento " & Trim(txtSerieComprobante.Text) & " - " & Trim(txtDocumentoComprobante.Text) & " tiene registrado Movimiento en Laboratorio" & Chr(13) & Chr(13) & "Fecha: " & oRsTmp.Fields!fecha & ",      N° Movimiento: " & oRsTmp.Fields!IdMovimiento, vbInformation, "Caja"
           BoletaTieneRegistradoLaboratorioImagenes = True
       End If
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
End Function

Private Sub txtTotal_KeyUp(KeyCode As Integer, Shift As Integer)
    'txtObservaciones.Text = DevolverConceptoPorDefecto
End Sub

'kike 2017
Private Sub txtTotal_LostFocus()
    If Not (Val(txtTotal.Text) > 0 And Val(txtTotal.Text) <= lnTotalBoleta) Then
       MsgBox "El PARCIAL debe ser mayor a CERO y menor a " & Trim(Str(lnTotalBoleta)), vbInformation, ""
       txtTotal.Text = lnTotalBoleta
    End If
End Sub

Sub CargaDatosParaPagoAutomatico()
    lnIdCaja = 0: lnIdGestionCaja = 0: lnIdTurno = 0
    If SIGHEntidades.Parametro378valorInt = 1 And mi_Opcion = sghAgregar Then
        Dim oRsTmp As New Recordset
        Set oRsTmp = mo_AdminCaja.CajaGestionSeleccionarXFiltroOrdenadoXFApertura("estadoLote='A' and idCajero=" & SIGHEntidades.Usuario)
        If oRsTmp.RecordCount > 0 Then
           oRsTmp.MoveFirst
           Do While Not oRsTmp.EOF
              If CDate(Format(oRsTmp!FechaApertura, SIGHEntidades.DevuelveFechaSoloFormato_DMY)) = Date Then
                lnIdCaja = oRsTmp!IdCaja
                lnIdGestionCaja = oRsTmp!IdGestionCaja
                lnIdTurno = oRsTmp!IdTurno
                Exit Do
              End If
              oRsTmp.MoveNext
           Loop
        End If
        oRsTmp.Close
        Set oRsTmp = Nothing
        If lnIdGestionCaja = 0 Then
           MsgBox "El parametro 378 está activo, para registrar automáticamente el pago en CAJA" & Chr(13) & _
                  "                            falta APERTURAR CAJA                            ", vbInformation, ""
           Me.btnAceptar.Visible = False
        End If
    End If
End Sub
