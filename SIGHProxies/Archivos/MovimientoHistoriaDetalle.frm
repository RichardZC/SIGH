VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form MovimientoHistoriaDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16170
   Icon            =   "MovimientoHistoriaDetalle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   16170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab 
      Height          =   5460
      Left            =   0
      TabIndex        =   36
      Top             =   1800
      Width           =   16125
      _ExtentX        =   28443
      _ExtentY        =   9631
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
      TabCaption(0)   =   "Historias"
      TabPicture(0)   =   "MovimientoHistoriaDetalle.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Datos complementarios"
      TabPicture(1)   =   "MovimientoHistoriaDetalle.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraMovimiento"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraMovimiento 
         Height          =   4830
         Left            =   -74880
         TabIndex        =   41
         Top             =   480
         Width           =   12030
         Begin VB.TextBox txtIdEmpleadoArchivo 
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
            Left            =   1590
            TabIndex        =   51
            Top             =   180
            Width           =   465
         End
         Begin VB.TextBox txtNombreEmpleadoArchivo 
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
            Left            =   2430
            TabIndex        =   50
            Top             =   180
            Width           =   4875
         End
         Begin VB.TextBox txtIdEmpleadoTransporte 
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
            Left            =   1590
            TabIndex        =   49
            Top             =   540
            Width           =   495
         End
         Begin VB.TextBox txtNombreEmpleadoTransporte 
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
            Left            =   2430
            TabIndex        =   48
            Top             =   540
            Width           =   4875
         End
         Begin VB.TextBox txtIdEmpleadoRecepcion 
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
            Left            =   1590
            TabIndex        =   47
            Top             =   930
            Width           =   495
         End
         Begin VB.TextBox txtNombreEmpleadoRecepcion 
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
            Left            =   2430
            TabIndex        =   46
            Top             =   930
            Width           =   4875
         End
         Begin VB.TextBox txtObservacion 
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
            Left            =   1590
            TabIndex        =   45
            Top             =   1290
            Width           =   5700
         End
         Begin VB.CommandButton btnBuscarRespArchivo 
            Caption         =   "..."
            Height          =   315
            Left            =   2100
            TabIndex        =   44
            Top             =   180
            Width           =   315
         End
         Begin VB.CommandButton btnBuscarRespTransporte 
            Caption         =   "..."
            Height          =   315
            Left            =   2100
            TabIndex        =   43
            Top             =   540
            Width           =   315
         End
         Begin VB.CommandButton btnBuscarRespRecepcion 
            Caption         =   "..."
            Height          =   315
            Left            =   2100
            TabIndex        =   42
            Top             =   930
            Width           =   315
         End
         Begin MSMask.MaskEdBox txtHoraMovimiento 
            Height          =   315
            Left            =   3120
            TabIndex        =   52
            Top             =   1605
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
         Begin MSMask.MaskEdBox txtFechaMovimiento 
            Height          =   315
            Left            =   1590
            TabIndex        =   53
            Top             =   1605
            Width           =   1380
            _ExtentX        =   2434
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
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Resp.Salida"
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
            TabIndex        =   58
            Top             =   225
            Width           =   900
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Resp.Transporte"
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
            TabIndex        =   57
            Top             =   585
            Width           =   1350
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Resp.Recepción"
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
            TabIndex        =   56
            Top             =   960
            Width           =   1290
         End
         Begin VB.Label Label6 
            Caption         =   "Observación"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   60
            TabIndex        =   55
            Top             =   1320
            Width           =   1065
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
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
            Height          =   210
            Left            =   60
            TabIndex        =   54
            Top             =   1665
            Width           =   1065
         End
      End
      Begin VB.Frame Frame1 
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
         Left            =   45
         TabIndex        =   37
         Top             =   330
         Width           =   15990
         Begin VB.CheckBox chkServiciosTodos 
            Caption         =   "Todos/Ninguno"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   60
            TabIndex        =   39
            Top             =   4800
            Width           =   1785
         End
         Begin SISGalenPlus.ucMensajeParpadeando ucMensajeParpadeando1 
            Height          =   300
            Left            =   3540
            TabIndex        =   38
            Top             =   4755
            Visible         =   0   'False
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   529
         End
         Begin UltraGrid.SSUltraGrid grdHistoriasSeleccionadas 
            Height          =   4575
            Left            =   60
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   180
            Width           =   15855
            _ExtentX        =   27966
            _ExtentY        =   8070
            _Version        =   131072
            GridFlags       =   17040384
            LayoutFlags     =   67108884
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Historias seleccionadas"
         End
      End
   End
   Begin VB.Frame fraFiltro 
      Caption         =   "Filtros"
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
      Left            =   3615
      TabIndex        =   23
      Top             =   30
      Width           =   12480
      Begin VB.OptionButton OptDevolverporServicio 
         Caption         =   "Devolver por Servicio"
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
         Left            =   105
         TabIndex        =   35
         Top             =   1290
         Width           =   2040
      End
      Begin VB.OptionButton OptDevolverHcXNroHistoria 
         Caption         =   "Devolver por HC"
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
         Left            =   105
         TabIndex        =   34
         Top             =   1005
         Width           =   1680
      End
      Begin VB.TextBox txtFichaFamiliar 
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
         Left            =   1170
         TabIndex        =   32
         Top             =   660
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Frame frmFiltro2 
         Height          =   1005
         Left            =   2265
         TabIndex        =   27
         Top             =   600
         Width           =   10110
         Begin VB.CheckBox chkTodosServ 
            Alignment       =   1  'Right Justify
            Caption         =   "Solo Servicios con Solicitud"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7500
            TabIndex        =   60
            Top             =   600
            Width           =   2520
         End
         Begin VB.ComboBox cmbTurnos 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   7410
            TabIndex        =   59
            Text            =   "cmbTurnos"
            ToolTipText     =   "Solo se muestran los servicios que corresponden al archivero"
            Top             =   120
            Width           =   2625
         End
         Begin VB.ComboBox cmbCondicionFechas 
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
            ItemData        =   "MovimientoHistoriaDetalle.frx":0D02
            Left            =   1455
            List            =   "MovimientoHistoriaDetalle.frx":0D15
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   570
            Width           =   1155
         End
         Begin VB.ComboBox cmbIdServicio 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1470
            Style           =   2  'Dropdown List
            TabIndex        =   5
            ToolTipText     =   "Solo se muestran los servicios que corresponden al archivero"
            Top             =   120
            Width           =   5895
         End
         Begin VB.ComboBox cmbFecha 
            BackColor       =   &H00E0E0E0&
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
            ItemData        =   "MovimientoHistoriaDetalle.frx":0D37
            Left            =   45
            List            =   "MovimientoHistoriaDetalle.frx":0D44
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   570
            Width           =   1395
         End
         Begin MSMask.MaskEdBox txtFechaDesde 
            Height          =   345
            Left            =   2595
            TabIndex        =   8
            Top             =   570
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtfechaHasta 
            Height          =   345
            Left            =   5205
            TabIndex        =   9
            Top             =   555
            Visible         =   0   'False
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
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
         Begin VB.Label lblServicioDestino 
            Caption         =   "Servicio destino"
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
            Left            =   45
            TabIndex        =   29
            Top             =   225
            Width           =   1365
         End
         Begin VB.Label lblHasta 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
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
            Left            =   4725
            TabIndex        =   28
            Top             =   615
            Visible         =   0   'False
            Width           =   450
         End
      End
      Begin VB.CommandButton btnBuscarPaciente 
         Caption         =   "..."
         Height          =   315
         Left            =   2250
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   225
         Width           =   315
      End
      Begin VB.TextBox txtIdHistoriaClinica 
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
         Left            =   1170
         TabIndex        =   2
         Top             =   240
         Width           =   1050
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   11010
         Picture         =   "MovimientoHistoriaDetalle.frx":0D74
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Haga click en este botón para filtrar las historias solicitadas"
         Top             =   300
         Width           =   1305
      End
      Begin VB.Label lblFichaFamiliar 
         AutoSize        =   -1  'True
         Caption         =   "Ficha Familiar"
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
         TabIndex        =   33
         Top             =   720
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label lblApellidos 
         Caption         =   "..."
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
         Left            =   2640
         TabIndex        =   26
         Top             =   270
         Width           =   4200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº Historia "
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
         TabIndex        =   25
         Top             =   300
         Width           =   915
      End
   End
   Begin VB.Frame fraDestino 
      Height          =   435
      Left            =   5310
      TabIndex        =   17
      Top             =   60
      Visible         =   0   'False
      Width           =   5205
      Begin VB.CommandButton btnAgregarDx 
         DisabledPicture =   "MovimientoHistoriaDetalle.frx":39BD
         DownPicture     =   "MovimientoHistoriaDetalle.frx":3DA6
         Height          =   315
         Left            =   3180
         Picture         =   "MovimientoHistoriaDetalle.frx":41B2
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   570
         Width           =   1005
      End
      Begin VB.CommandButton btnQuitarDx 
         DisabledPicture =   "MovimientoHistoriaDetalle.frx":45BE
         DownPicture     =   "MovimientoHistoriaDetalle.frx":4949
         Height          =   315
         Left            =   4260
         Picture         =   "MovimientoHistoriaDetalle.frx":4CDC
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   570
         Width           =   1005
      End
      Begin VB.TextBox txtNombreServicioDestino 
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
         Left            =   3180
         TabIndex        =   19
         Top             =   180
         Width           =   5115
      End
      Begin VB.TextBox txtIdServicioDestino 
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
         Left            =   1590
         TabIndex        =   18
         Top             =   195
         Width           =   1155
      End
      Begin VB.Label Label5 
         Caption         =   "Servicio destino"
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
         Left            =   180
         TabIndex        =   22
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.CommandButton btnListarMovimientosAsoc 
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
      Height          =   405
      Left            =   1680
      TabIndex        =   10
      Top             =   6960
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   15
      TabIndex        =   14
      Top             =   7200
      Width           =   16110
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "MovimientoHistoriaDetalle.frx":506D
         DownPicture     =   "MovimientoHistoriaDetalle.frx":5531
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
         Left            =   8175
         Picture         =   "MovimientoHistoriaDetalle.frx":5A1D
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "MovimientoHistoriaDetalle.frx":5F09
         DownPicture     =   "MovimientoHistoriaDetalle.frx":6369
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
         Left            =   6630
         Picture         =   "MovimientoHistoriaDetalle.frx":67DE
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1365
      End
      Begin SISGalenPlus.XP_ProgressBar progressRpt 
         Height          =   300
         Left            =   90
         TabIndex        =   15
         Top             =   300
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
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
   End
   Begin VB.Frame fraPaciente 
      Height          =   1665
      Left            =   30
      TabIndex        =   13
      Top             =   30
      Width           =   3585
      Begin VB.CommandButton cmdAgregaPaciente 
         Caption         =   "Agregar Paciente"
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
         TabIndex        =   31
         Top             =   930
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.ComboBox cmbIdServicioOrigen 
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
         Left            =   1095
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Solo se muestran los servicios que corresponden al archivero"
         Top             =   1230
         Visible         =   0   'False
         Width           =   2475
      End
      Begin VB.ComboBox cmbIdMotivo 
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
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Seleccionar el motivo por el cual se va a realizar el movimiento"
         Top             =   180
         Width           =   2445
      End
      Begin VB.Label lblServicioOrigen 
         AutoSize        =   -1  'True
         Caption         =   "Serv.Origen"
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
         TabIndex        =   30
         Top             =   1290
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label lblArchivero 
         Caption         =   "Archivero: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         TabIndex        =   24
         Top             =   630
         Width           =   3270
      End
      Begin VB.Label Label3 
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
         Height          =   240
         Left            =   90
         TabIndex        =   16
         Top             =   210
         Width           =   645
      End
   End
End
Attribute VB_Name = "MovimientoHistoriaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Movimiento de Historia
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_MovimientosHistoriaClinica As New DOMovimientoHistoriaClinica
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_AdminComun As New SIGHNegocios.ReglasComunes
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim mrs_HistoriasPorMover As New ADODB.Recordset
Dim mo_Movimientos As New Collection
Dim mo_cmbIdMotivo As New sighentidades.ListaDespleglable
Dim mo_cmbIdServicio As New sighentidades.ListaDespleglable
Dim mo_cmbIdServiciOrigen As New sighentidades.ListaDespleglable
Dim ml_IdMovimiento As Long
Dim ml_IdGrupoMovimiento As Long
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_Movimiento As DOMovimientoHistoriaClinica
Dim ml_idPaciente As Long
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim lcHCyPaciente As String
Dim lnUsuarioFiltroCombo As Long
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lcDefaultHistoriaFicha As String
Dim ml_idPacienteSeleccionado As Long
Dim lcFiltroServicios As String

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Property Let idPacienteSeleccionado(lValue As Long)
    ml_idPacienteSeleccionado = lValue
End Property

Sub CargarComboBoxes()
Dim sSQL As String
Dim sMensaje As String

       mo_cmbIdMotivo.BoundColumn = "IdMotivo"
       mo_cmbIdMotivo.ListField = "DescripcionLarga"
       Set mo_cmbIdMotivo.RowSource = mo_AdminArchivoClinico.MotivosMovimientoHistoriaSeleccionarTodos()
'       mo_cmbIdMotivo.BoundText = "1" 'Frank
       
       sMensaje = mo_AdminArchivoClinico.MensajeError
       If sMensaje <> "" Then
           MsgBox mo_AdminArchivoClinico.MensajeError, vbInformation, Me.Caption
       End If


        

End Sub
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
Property Let IdMovimiento(lValue As Long)
   ml_IdMovimiento = lValue
End Property
Property Get IdMovimiento() As Long
   IdMovimiento = ml_IdMovimiento
End Property

Private Sub btnBuscar_Click()
10        On Error GoTo ErrBuscar
      '    If Trim(Me.cmbIdServicio.Text) <> "" Then
      '       Me.lblApellidos.Caption = "..."
      '    End If
          'mgaray201410c
20        If Me.lblApellidos.Caption = "" And OptDevolverHcXNroHistoria.Value = True Then
30            MsgBox "Tiene que registrar la Historia", vbInformation, Me.Caption
40            Exit Sub
50        End If
60        If mo_cmbIdMotivo.BoundText = "" Then
70            MsgBox "Ingrese el motivo del movimiento", vbInformation, Me.Caption
80            Exit Sub
90        End If
100       ucMensajeParpadeando1.MensajeDeTexto = ""
110       ucMensajeParpadeando1.Visible = False
          
120       Select Case Val(mo_cmbIdMotivo.BoundText)
          Case 9
130               If OptDevolverHcXNroHistoria.Value = True Then
140                   If lblApellidos.Caption = "" Or lblApellidos.Caption = "..." Then
150                       MsgBox "Ingrese el N° Historia a Devolver", vbInformation, Me.Caption
160                       Exit Sub
170                   End If
180               Else
190                   If OptDevolverporServicio.Value = True Then
200                       If cmbIdServicio.Text = "" Then
210                          MsgBox "Debe elegir el Servicio", vbInformation, "Mensaje"
220                          Me.cmbIdServicio.SetFocus
230                          Exit Sub
240                       End If
250                   End If
260               End If
270       Case 4, 5, 6, 8, 10
                 '***************** GalenHos v.3.0 (inicio)*****************
280              If cmbIdServicio.Text = "" Then
290                 MsgBox "Debe elegir el Servicio", vbInformation, "Mensaje"
300                 Me.cmbIdServicio.SetFocus
310                 Exit Sub
320              End If
330              If lblApellidos.Caption = "..." Or lblApellidos.Caption = "" Then
340                 MsgBox "Debe registrar el Nro Historia Clinica", vbInformation, "Mensaje"
350                 Exit Sub
360              End If
370              If ml_idPaciente = 0 Then
380                 MsgBox "Debe registrar el Nro Historia Clinica", vbInformation, "Mensaje"
390                 Exit Sub
400              End If
                 Dim oDOHistoriaSolicitada As New DOHistoriaSolicitada
                 Dim oHistoriasSolicitadas As New HistoriasSolicitadas
                 Dim oConexion As New Connection
410              oConexion.Open sighentidades.CadenaConexion
420              oConexion.CursorLocation = adUseClient
430              Set oHistoriasSolicitadas.Conexion = oConexion
440              oDOHistoriaSolicitada.FechaRequerida = Me.txtFechaDesde.Text
450              oDOHistoriaSolicitada.FechaSolicitud = Me.txtFechaDesde.Text
460              oDOHistoriaSolicitada.HoraRequerida = Format(Now, sighentidades.DevuelveHoraSoloFormato_HM)
470              oDOHistoriaSolicitada.HoraSolicitud = Format(Now, sighentidades.DevuelveHoraSoloFormato_HM)
480              oDOHistoriaSolicitada.IdEmpleadoSolicita = ml_idUsuario
490              oDOHistoriaSolicitada.IdMotivo = Val(mo_cmbIdMotivo.BoundText)
500              oDOHistoriaSolicitada.idPaciente = ml_idPaciente
510              oDOHistoriaSolicitada.IdServicio = Val(mo_cmbIdServicio.BoundText)
520              oDOHistoriaSolicitada.IdUsuarioAuditoria = ml_idUsuario
530              If Not oHistoriasSolicitadas.Insertar(oDOHistoriaSolicitada) Then
540                 MsgBox oHistoriasSolicitadas.MensajeError: GoTo ErrBuscar
550              End If
560              oConexion.Close
570              Set oDOHistoriaSolicitada = Nothing
580              Set oHistoriasSolicitadas = Nothing
590              Set oConexion = Nothing
                 '***************** GalenHos v.3.0 (fin)*****************
600       Case 7
                 '***************** GalenHos v.3.0 (inicio)*****************
610              If cmbIdServicioOrigen.Text = "" Then
620                 MsgBox "Debe elegir el Servicio ORIGEN del Hospital", vbInformation, "Mensaje"
630                 Exit Sub
640              End If
650              If cmbIdServicio.Text = "" Then
660                 MsgBox "Debe elegir el Servicio DESTINO del Hospital", vbInformation, "Mensaje"
670                 Exit Sub
680              End If
690              If lblApellidos.Caption = "" Then
700                 MsgBox "Debe registrar el Nro Historia Clinica", vbInformation, "Mensaje"
710                 Exit Sub
720              End If
730               mrs_HistoriasPorMover.AddNew
740               mrs_HistoriasPorMover!seleccionar = True
750               mrs_HistoriasPorMover!IdHistoriaSolicitada = 0
760               mrs_HistoriasPorMover!idPaciente = ml_idPaciente
770               mrs_HistoriasPorMover!HistoriaClinica = Me.txtIdHistoriaClinica.Text
780               mrs_HistoriasPorMover!Nombres = lblApellidos.Caption
790               mrs_HistoriasPorMover!FechaSolicitud = Date
800               mrs_HistoriasPorMover!FechaRequerida = Date
810               mrs_HistoriasPorMover!NroFolios = 0
820               mrs_HistoriasPorMover!idServicioDestino = Val(mo_cmbIdServicio.BoundText)
830               mrs_HistoriasPorMover!nombreServicioDestino = cmbIdServicio.Text
840               mrs_HistoriasPorMover!IdServicioOrigen = Val(mo_cmbIdServiciOrigen.BoundText)
850               mrs_HistoriasPorMover!nombreServicioOrigen = cmbIdServicioOrigen.Text
860               mrs_HistoriasPorMover!idTipoHistoria = 0
870               mrs_HistoriasPorMover!IdMovimientoHistoria = 0
880               mrs_HistoriasPorMover!IdEstadoregistro = "A"
890               mrs_HistoriasPorMover!FormaPago = 0
900               mrs_HistoriasPorMover!PagoCita = ""
910               mrs_HistoriasPorMover!idAtencion = 0
920               mrs_HistoriasPorMover!SeDaraSalida = True
930               mrs_HistoriasPorMover.Update
940               Exit Sub
                  '***************** GalenHos v.3.0 (fin)*****************
950       Case Else
                 
960               If Me.cmbCondicionFechas.ListIndex <> 0 Then
970                   If Me.txtFechaDesde = sighentidades.FECHA_VACIA_DMY Then
980                       MsgBox "Ingrese la Fecha Desde", vbInformation, Me.Caption
990                   End If
1000              End If
                  
1010              If Me.cmbCondicionFechas.ListIndex = 4 Then
1020                  If Me.txtFechaHasta = sighentidades.FECHA_VACIA_DMY Then
1030                      MsgBox "Ingrese la Fecha Hasta", vbInformation, Me.Caption
1040                  End If
1050              Else
1060                  Me.txtFechaHasta = sighentidades.FECHA_VACIA_DMY
1070              End If
          
          
1080      End Select

          Dim sOperador As String
1090      sOperador = Trim(cmbCondicionFechas.List(cmbCondicionFechas.ListIndex))

1100      LimpiarGrilla

          Dim rsSolicitudes As New Recordset
          
1110      Select Case Val(mo_cmbIdMotivo.BoundText)
          Case 9
          ' yamill
1120          If OptDevolverporServicio.Value = False Then
1130              Set rsSolicitudes = mo_AdminArchivoClinico.MovimientosHistoriasClinicasParaDevolverPorNroHistoria(Val(mo_cmbIdServicio.BoundText), Me.txtFechaDesde.Text, ml_idUsuario, Val(txtIdHistoriaClinica), mo_AdminComun.ParametrosIdServicioArchivoClinico(), ml_idPaciente)
1140          Else
1150              Set rsSolicitudes = mo_AdminArchivoClinico.MovimientosHistoriasClinicasParaDevolverPorServicio(Val(mo_cmbIdServicio.BoundText), Me.txtFechaDesde.Text, ml_idUsuario, Val(txtIdHistoriaClinica), mo_AdminComun.ParametrosIdServicioArchivoClinico(), 0, cmbFecha.ListIndex, sOperador, Me.txtFechaDesde.Text, IIf(Me.txtFechaHasta = sighentidades.FECHA_VACIA_DMY, "", Me.txtFechaHasta), Me.txtFechaHasta.Visible)
1160          End If
1170      Case Else
1180          Set rsSolicitudes = mo_AdminArchivoClinico.HistoriasSolicitadasSeleccionarPorArchivero(ml_idUsuario, Val(mo_cmbIdMotivo.BoundText), Val(mo_cmbIdServicio.BoundText), sOperador, Me.txtFechaDesde.Text, IIf(Me.txtFechaHasta = sighentidades.FECHA_VACIA_DMY, "", Me.txtFechaHasta), Val(txtIdHistoriaClinica.Text), IIf(cmbFecha.ListIndex = 1, True, False), mo_AdminComun.ParametrosIdServicioArchivoClinico())
1190          If Val(mo_cmbIdMotivo.BoundText) >= 1 And Val(mo_cmbIdMotivo.BoundText) <= 3 Then
1200             rsSolicitudes.Filter = "idAtencion>0"
1210          End If
1220      End Select
          
1230      LlenarGrilladeHistoriasSeleccionadas rsSolicitudes, Val(mo_cmbIdMotivo.BoundText)
1240      Exit Sub
ErrBuscar:
          MsgBox Err.Number & " " & Err.Description & _
                sighentidades.DevuelveFuenteDeLineaDelError(Erl(), "Private Sub btnBuscar_Click", "movimientoHistoriaDetalle.frm")   'debb-02/05/2016


End Sub

Sub LimpiarGrilla()

    
        If mrs_HistoriasPorMover Is Nothing Then
            Exit Sub
        End If
        If mrs_HistoriasPorMover.RecordCount > 0 Then
            mrs_HistoriasPorMover.MoveFirst
            Do While Not mrs_HistoriasPorMover.EOF
                mrs_HistoriasPorMover.Delete
                mrs_HistoriasPorMover.Update
                mrs_HistoriasPorMover.MoveNext
            Loop
        End If
End Sub

Private Sub btnBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
   AdministrarKeyPreview KeyCode

End Sub

Private Sub btnBuscarPaciente_Click()
Dim oBusqueda As New SIGHNegocios.BuscaPacientes
Dim oDOPaciente As New doPaciente
Dim oConexion As New Connection
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
    oBusqueda.TipoFiltro = sghFiltrarConHistoriasDefinitivas
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oConexion)
        If Not oDOPaciente Is Nothing Then
            ml_idPaciente = oDOPaciente.idPaciente
            Me.txtIdHistoriaClinica.Text = oDOPaciente.NroHistoriaClinica
            lblApellidos.Caption = Trim(oDOPaciente.ApellidoPaterno) + " " + Trim(oDOPaciente.ApellidoMaterno) + " " + Trim(oDOPaciente.PrimerNombre)
            btnBuscar.SetFocus
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
End Sub

Private Sub btnBuscarRespArchivo_Click()
    CompletarDatosResponsable Me.txtIdEmpleadoArchivo, Me.txtNombreEmpleadoArchivo
End Sub

Private Sub btnBuscarRespRecepcion_Click()
    CompletarDatosResponsable Me.txtIdEmpleadoRecepcion, Me.txtNombreEmpleadoRecepcion
End Sub

Private Sub btnBuscarRespTransporte_Click()
    CompletarDatosResponsable Me.txtIdEmpleadoTransporte, Me.txtNombreEmpleadoTransporte
End Sub



Private Sub btnListarMovimientosAsoc_Click()
'Dim oRptMovimiento As New RptMovimientoHistorias
Dim oRptMovimiento As New SIGHReportes.clMovimientoHist

    oRptMovimiento.IdGrupoMovimiento = ml_IdGrupoMovimiento
    'Set oRptMovimiento.progressRpt = Me.progressRpt
    oRptMovimiento.CrearReporteMovimientoHistoria Me.hwnd
    
End Sub

Private Sub chkServiciosTodos_Click()
    If mrs_HistoriasPorMover.RecordCount > 0 Then
        mrs_HistoriasPorMover.MoveFirst
        Do While Not mrs_HistoriasPorMover.EOF
            If chkServiciosTodos.Value = 1 Then
               mrs_HistoriasPorMover.Fields!seleccionar = 1
            Else
               mrs_HistoriasPorMover.Fields!seleccionar = 0
            End If
            mrs_HistoriasPorMover.Update
            mrs_HistoriasPorMover.MoveNext
        Loop
    End If
End Sub

Private Sub chkServiciosTodos_KeyDown(KeyCode As Integer, Shift As Integer)
   AdministrarKeyPreview KeyCode

End Sub

Private Sub chkTodosServ_Click()
        ServiciosCargar
End Sub

Private Sub cmbCondicionFechas_Click()
    If Me.cmbCondicionFechas.ListIndex = 4 Then
        Me.lblHasta.Visible = True
        Me.txtFechaHasta.Visible = True
        Me.txtFechaHasta.Text = Format(Date, sighentidades.DevuelveFechaSoloFormato_DMY)
    Else
        Me.lblHasta.Visible = False
        Me.txtFechaHasta.Visible = False
        Me.txtFechaHasta.Text = sighentidades.FECHA_VACIA_DMY
    End If
    '
    If Me.cmbCondicionFechas.ListIndex = 1 Or Me.cmbCondicionFechas.ListIndex = 2 Or Me.cmbCondicionFechas.ListIndex = 3 Then
       chkTodosServ.Visible = True
    Else
       chkTodosServ.Visible = False
    End If
    ServiciosCargar
    '
End Sub

Sub ServiciosCargar()
    Dim lcIdServicioActual As String
    If Val(mo_cmbIdServicio.BoundText) > 0 Then
       lcIdServicioActual = mo_cmbIdServicio.BoundText
    End If
    mo_cmbIdServicio.BoundColumn = "IdServicio"
    mo_cmbIdServicio.ListField = "DescripcionLarga"
    If Me.chkTodosServ.Value = 0 Or Me.chkTodosServ.Visible = False Then
       Set mo_cmbIdServicio.RowSource = mo_AdminArchivoClinico.ServiciosPorArchiveroTipoServicio(lnUsuarioFiltroCombo, lcFiltroServicios)
    Else
       Set mo_cmbIdServicio.RowSource = ServiciosSoloConSolicitud
    End If
    If Val(lcIdServicioActual) > 0 Then
       On Error Resume Next
       mo_cmbIdServicio.BoundText = lcIdServicioActual
    End If
End Sub

Function ServiciosSoloConSolicitud() As Recordset
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim oRsTmp3 As New Recordset
    Dim oConexion As New ADODB.Connection
    Dim lcFiltro As String
    Dim lcHoraIniX As String, lcHoraFinX As String
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    With oRsTmp2
          .Fields.Append "IdServicio", adInteger, 4, adFldIsNullable
          .Fields.Append "DescripcionLarga", adVarChar, 200, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
    Set oRsTmp1 = mo_AdminArchivoClinico.ServiciosPorArchiveroTipoServicio(lnUsuarioFiltroCombo, lcFiltroServicios)
    If oRsTmp1.RecordCount > 0 Then
       If Trim(cmbTurnos.Text) = "" Then
            lcHoraIniX = "00:01"
            lcHoraFinX = "23:59"
       Else
            lcHoraIniX = Left(cmbTurnos.Text, 5)
            lcHoraFinX = Mid(cmbTurnos.Text, 7, 5)
       End If
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
          lcFiltro = ""
          Select Case cmbCondicionFechas.ListIndex
          Case 1   '=
                lcFiltro = " HistoriasSolicitadas.idServicio=" & oRsTmp1!IdServicio & _
                           " and HistoriasSolicitadas.fechaRequerida Between (CONVERT(DATETIME,'" & _
                           txtFechaDesde.Text & "',103)) and (CONVERT(DATETIME,'" & txtFechaDesde.Text & _
                           "',103)) and (HistoriasSolicitadas.HoraRequerida>='" & lcHoraIniX & "' and " & _
                           " HistoriasSolicitadas.HoraRequerida<='" & lcHoraFinX & "')"
          Case 2   '>=
                lcFiltro = " HistoriasSolicitadas.idServicio=" & oRsTmp1!IdServicio & _
                           " and HistoriasSolicitadas.fechaRequerida Between (CONVERT(DATETIME,'" & _
                           txtFechaDesde.Text & "',103)) and (CONVERT(DATETIME,'01/01/2099" & _
                           "',103)) and (HistoriasSolicitadas.HoraRequerida>='" & lcHoraIniX & "' and " & _
                           " HistoriasSolicitadas.HoraRequerida<='" & lcHoraFinX & "')"
          Case 3   '<=
                lcFiltro = " HistoriasSolicitadas.idServicio=" & oRsTmp1!IdServicio & _
                           " and HistoriasSolicitadas.fechaRequerida Between (CONVERT(DATETIME,'" & _
                            "01/01/1990',103)) and (CONVERT(DATETIME,'" & txtFechaDesde.Text & _
                           "',103)) and (HistoriasSolicitadas.HoraRequerida>='" & lcHoraIniX & "' and " & _
                           " HistoriasSolicitadas.HoraRequerida<='" & lcHoraFinX & "')"
          End Select
          Set oRsTmp3 = mo_AdminArchivoClinico.HistoriasSolicitadasSegunFiltro(lcFiltro, oConexion)
          If oRsTmp3.RecordCount > 0 Then
                oRsTmp2.AddNew
                oRsTmp2.Fields!IdServicio = oRsTmp1!IdServicio
                oRsTmp2.Fields!DescripcionLarga = oRsTmp1!DescripcionLarga
                oRsTmp2.Update
          End If
          oRsTmp3.Close
          oRsTmp1.MoveNext
       Loop
    End If
    oConexion.Close
    Set ServiciosSoloConSolicitud = oRsTmp2
    Set oRsTmp1 = Nothing
    Set oRsTmp2 = Nothing
    Set oConexion = Nothing
End Function

Private Sub cmbIdMotivo_Click()
    
    Me.txtIdServicioDestino.Tag = ""
    Me.txtIdServicioDestino.Text = ""
    Me.txtNombreServicioDestino = ""
    
    fraDestino.Visible = False
    '
    LimpiarGrilla
    Select Case mi_Opcion
    Case sghAgregar
        fraFiltro.Visible = True
        Me.cmbCondicionFechas.ListIndex = 1
        txtFechaDesde.Text = Format(Date, sighentidades.DevuelveFechaSoloFormato_DMY)
        txtFechaHasta.Text = Format(Date, sighentidades.DevuelveFechaSoloFormato_DMY)
    Case Else
        fraFiltro.Visible = False
    End Select
    frmFiltro2.Visible = True
    Me.cmbCondicionFechas.Visible = True
    cmbFecha.Visible = True
    txtFechaDesde.Visible = True
    cmbIdServicioOrigen.Visible = False
    lblServicioOrigen.Visible = False
    Select Case Val(mo_cmbIdMotivo.BoundText)
    Case 9
        lblServicioDestino.Caption = "Servicio"
        OptDevolverHcXNroHistoria.Visible = True
        OptDevolverporServicio.Visible = True
        OptDevolverHcXNroHistoria.Value = 1
        mostrarbusquedaparadevolucion
    Case Else
        '***************** GalenHos v.3.0 (inicio)*****************
        lblServicioDestino.Caption = "Servicio destino"
        frmFiltro2.Visible = True
        Me.cmbCondicionFechas.Enabled = True
        cmbIdServicio.Enabled = True
        cmbFecha.ListIndex = 2
        Label1.Visible = True
        txtIdHistoriaClinica.Visible = True
        lblFichaFamiliar.Visible = False
        txtFichaFamiliar.Visible = False
        If lcBuscaParametro.SeleccionaFilaParametro(282) = "S" Then
           Me.lblFichaFamiliar.Visible = True
           Me.txtFichaFamiliar.Visible = True
        End If
        btnBuscarPaciente.Visible = True
        lblApellidos.Visible = True
        
        OptDevolverHcXNroHistoria.Visible = False
        OptDevolverporServicio.Visible = False
        
        Me.cmbCondicionFechas.Enabled = True
        cmbIdServicio.Enabled = True
        cmbFecha.ListIndex = 1

        Select Case Val(mo_cmbIdMotivo.BoundText)
        
        Case 1      'CE
            lcFiltroServicios = "(1)"
            ServiciosCargar
'            mo_cmbIdServicio.BoundColumn = "IdServicio"
'            mo_cmbIdServicio.ListField = "DescripcionLarga"
'            Set mo_cmbIdServicio.RowSource = mo_AdminArchivoClinico.ServiciosPorArchiveroTipoServicio(lnUsuarioFiltroCombo, "(1)")
            Me.cmbCondicionFechas.ListIndex = 1
        Case 2      'Hospitalizacion
            lcFiltroServicios = "(3)"
            ServiciosCargar
'            mo_cmbIdServicio.BoundColumn = "IdServicio"
'            mo_cmbIdServicio.ListField = "DescripcionLarga"
'            Set mo_cmbIdServicio.RowSource = mo_AdminArchivoClinico.ServiciosPorArchiveroTipoServicio(lnUsuarioFiltroCombo, "(3)")
            Me.cmbCondicionFechas.ListIndex = 4
            txtFechaDesde.Text = DateAdd("m", -2, Date)
            txtFechaHasta.Text = Date
        Case 3      'Emergencia
            lcFiltroServicios = "(2,4)"
            ServiciosCargar
'            mo_cmbIdServicio.BoundColumn = "IdServicio"
'            mo_cmbIdServicio.ListField = "DescripcionLarga"
'            Set mo_cmbIdServicio.RowSource = mo_AdminArchivoClinico.ServiciosPorArchiveroTipoServicio(lnUsuarioFiltroCombo, "(2,4)")
            Me.cmbCondicionFechas.ListIndex = 4
            txtFechaDesde.Text = DateAdd("m", -2, Date)
            txtFechaHasta.Text = Date
        Case Else
            lcFiltroServicios = ""
            ServiciosCargar
'            mo_cmbIdServicio.BoundColumn = "IdServicio"
'            mo_cmbIdServicio.ListField = "DescripcionLarga"
'            Set mo_cmbIdServicio.RowSource = mo_AdminArchivoClinico.ServiciosPorArchiveroTipoServicio(lnUsuarioFiltroCombo, "")
            Me.cmbCondicionFechas.Visible = False
            cmbFecha.Visible = False
            txtFechaDesde.Visible = False
            If Val(mo_cmbIdMotivo.BoundText) = 7 Then
                mo_cmbIdServiciOrigen.BoundColumn = "IdServicio"
                mo_cmbIdServiciOrigen.ListField = "DescripcionLarga"
                Set mo_cmbIdServiciOrigen.RowSource = mo_AdminArchivoClinico.ServiciosPorArchiveroTipoServicio(lnUsuarioFiltroCombo, "")
                cmbIdServicioOrigen.Visible = True
                lblServicioOrigen.Visible = True
            End If
        End Select
        '***************** GalenHos v.3.0 (fin)*****************
    End Select

End Sub

'debb-12/04/2016
Sub LlenarGrilladeHistoriasSeleccionadas(rsSolicitudes As Recordset, idMotivoMovimiento As Long)
10            On Error GoTo ErrLlenarGrilla     'debb-02/05/2016
              Dim oRsCitaPagada As New ADODB.Recordset
              Dim oRsHCtieneMov As New ADODB.Recordset
              Dim oRsPaquete As New Recordset
              Dim lcSql As String, lnSolicitudes As Long
              Dim lcFormaPago As String
              Dim lnIdNroHistoria As Long
              Dim lbContinua As Boolean
              Dim lbSeDaraSalida As Boolean
              Dim lnIdOrigen As Long, lcOrigen As String
              Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
              Dim oConexion As New Connection
              Dim lcEsNuevaHistoria As String
              'mgaray201410c
              Dim lcMotivoNoSalidaHistoria As String
              Dim lbEsServicioCostoCero As Boolean
              Dim oRsHistoriasSinSalida As ADODB.Recordset
              Dim lcHoraIniX As String, lcHoraFinX As String, lcHoraRequeridaX As String
              Dim lnNroCitasEnOtroServicio As Long, lcObservaciones As String
              Dim lcFallecioPaciente As String, lnIdPaciente1 As Long
20            Set oRsHistoriasSinSalida = inicializarRsHistoriasClinicasRestringidas()
              
30            oConexion.CommandTimeout = 300
40            oConexion.CursorLocation = adUseClient
50            oConexion.Open sighentidades.CadenaConexion
              
              
              
60            lnSolicitudes = rsSolicitudes.RecordCount
70            If lnSolicitudes > 0 Then
80                ms_MensajeError = ""
90                rsSolicitudes.MoveFirst
100               Do While Not rsSolicitudes.EOF
110                   lnIdPaciente1 = rsSolicitudes!idPaciente
120                   lbEsServicioCostoCero = False
130                  lcSql = " ": lcFormaPago = "": lbSeDaraSalida = True
140                  lcOrigen = rsSolicitudes!nombreServicioOrigen
150                  lnIdOrigen = rsSolicitudes!IdServicioOrigen
160                  lbContinua = True
                     
                     '***************** GalenHos v.3.0 (inicio)*****************
                     'Chequea si el ultimo Movimiento corresponde al MOTIVO
170                  Set oRsHCtieneMov = mo_AdminArchivoClinico.MovimientosHistoriaClinicaSeleccionaUltimoMovimientoPorPaciente(rsSolicitudes!idPaciente)
180                  If oRsHCtieneMov.RecordCount > 0 Then
190                     Set oRsCitaPagada = mo_AdminArchivoClinico.HistoriasSolicitadasXidentificador(oRsHCtieneMov.Fields!IdMovimiento)
200                     If oRsCitaPagada.RecordCount > 0 Then
210                          lcSql = " (F.Requer: " & oRsCitaPagada.Fields!FechaRequerida & ")"
220                     Else
230                          lcSql = ""
240                     End If
250                     oRsCitaPagada.Close
260                     If idMotivoMovimiento = 9 Then
270                        If oRsHCtieneMov.Fields!IdMotivo = 9 Then
280                           lbContinua = False
290                           ms_MensajeError = ms_MensajeError & "La HC: " & Trim(rsSolicitudes!NroHistoriaClinica) & " ya RETORNO el " & oRsHCtieneMov!FechaMovimiento & lcSql & " (" & Trim(oRsHCtieneMov!nombreServicioOrigen) & ")" & Chr(13)
300                           ucMensajeParpadeando1.MensajeDeTexto = "La HC: " & Trim(rsSolicitudes!NroHistoriaClinica) & " ya tubo RETORNO el " & oRsHCtieneMov!FechaMovimiento & lcSql & " (" & oRsHCtieneMov!nombreServicioOrigen & ")"
310                           ucMensajeParpadeando1.Visible = True
320                        Else
330                           lcOrigen = oRsHCtieneMov!nombreServicioDestino
340                           lnIdOrigen = oRsHCtieneMov!idServicioDestino
350                        End If
360                     Else
370                        If oRsHCtieneMov.Fields!IdMotivo <> 9 Then
380                           ms_MensajeError = ms_MensajeError & "La HC: " & Trim(rsSolicitudes!NroHistoriaClinica) & " ya SALIO el " & oRsHCtieneMov!FechaMovimiento & lcSql & " (" & Trim(oRsHCtieneMov!nombreServicioDestino) & ")" & Chr(13)
390                           ucMensajeParpadeando1.MensajeDeTexto = "La HC: " & Trim(rsSolicitudes!NroHistoriaClinica) & " ya tubo SALIDA el " & oRsHCtieneMov!FechaMovimiento & lcSql & " (" & Trim(oRsHCtieneMov!nombreServicioDestino) & ")"
400                           ucMensajeParpadeando1.Visible = True
410                           lbContinua = False
420                        End If
430                     End If
440                  End If
450                  oRsHCtieneMov.Close
                     '***************** GalenHos v.3.0 (fin)*****************
                     '
460                  If Not IsNull(rsSolicitudes!idAtencion) And lbContinua = True And idMotivoMovimiento <> 9 Then
470                        Set oRsCitaPagada = mo_ReglasFacturacion.DevuelveSiPagoConsultaMedicaEnCaja(rsSolicitudes!idAtencion, mo_cmbIdMotivo.BoundText)
480                        lcSql = " "
490                        If oRsCitaPagada.RecordCount > 0 Then
500                           If oRsCitaPagada.Fields!idestadofacturacion = 4 Then
510                              lcSql = "Si"
520                           Else
530                              Set oRsPaquete = mo_ReglasFacturacion.FacturacionPaquetesSeleccionarPorFiltro(" idPuntoCarga=6 and AtencionId=" & oRsCitaPagada.Fields!idCuentaAtencion)
540                              If oRsPaquete.RecordCount > 0 Then
550                                   lcSql = "Si"
560                              Else
570                                   If mo_cmbIdMotivo.BoundText = "1" And oRsCitaPagada.Fields!IdFormaPago = 1 Then
580                                       If mo_AdminAdmision.EsServicioCostoCero(oRsCitaPagada.Fields!IdServicioIngreso) = True Then
                                              'mgaray201410c
590                                           lcMotivoNoSalidaHistoria = "No pagó"
600                                           lbEsServicioCostoCero = True
      '                                       If MsgBox("Destino:" + rsSolicitudes!NombreServicioDestino + Chr(13) + Chr(13) + "No pagó     ¿ desea darle SALIDA de todas maneras ?  ", vbQuestion + vbYesNo, "Servicio") = vbNo Then
610                                                ms_MensajeError = ms_MensajeError & "La HC: " & Trim(rsSolicitudes!NroHistoriaClinica) & " no podrá darle SALIDA porque no Pagó" & Chr(13)
620                                                lbSeDaraSalida = False
      '                                       End If
630                                       Else
                                              'mgaray201410c
640                                           lcMotivoNoSalidaHistoria = "No pagó"
650                                          ms_MensajeError = ms_MensajeError & "La HC: " & Trim(rsSolicitudes!NroHistoriaClinica) & " no podrá darle SALIDA porque no Pagó" & Chr(13)
660                                          lbSeDaraSalida = False
670                                       End If
680                                   End If
690                              End If
700                              oRsPaquete.Close
710                           End If
720                           lcFormaPago = oRsCitaPagada!Descripcion
730                        Else
740                           oRsCitaPagada.Close
750                           Set oRsCitaPagada = mo_AdminArchivoClinico.HistoriasPagoCitaDescripcionTarifa(rsSolicitudes!idAtencion, mo_cmbIdMotivo.BoundText)
760                           lcSql = " "
770                           If oRsCitaPagada.RecordCount > 0 Then
780                              lcFormaPago = oRsCitaPagada!Descripcion
790                              If mo_cmbIdMotivo.BoundText = "1" And oRsCitaPagada.Fields!IdFormaPago = 1 Then
                                      'mgaray201410c
800                                   lcMotivoNoSalidaHistoria = "No pagó"
810                                  ms_MensajeError = ms_MensajeError & "La HC: " & Trim(rsSolicitudes!NroHistoriaClinica) & " no podrá darle SALIDA porque no Pagó" & Chr(13)
820                                  lbSeDaraSalida = False
830                              End If
840                           Else
850                              lbContinua = False
860                           End If
870                        End If
880                        oRsCitaPagada.Close
890                   End If
900                   If lblApellidos.Caption = "" And Trim(cmbTurnos.Text) <> "" And lbContinua = True Then
910                      lcHoraIniX = Left(cmbTurnos.Text, 5)
920                      lcHoraFinX = Mid(cmbTurnos.Text, 7, 5)
930                      If sighentidades.EsHora(lcHoraIniX) = True And sighentidades.EsHora(lcHoraFinX) Then
940                         lcHoraRequeridaX = Format(rsSolicitudes!FechaRequerida, sighentidades.DevuelveHoraSoloFormato_HM)
950                         If Not (lcHoraRequeridaX >= lcHoraIniX And lcHoraRequeridaX <= lcHoraFinX) Then
960                            lbContinua = False
970                         End If
980                      End If
990                   End If
1000                  If Trim(cmbCondicionFechas.Text) = "=" And lbContinua = True Then
1010  lbContinua = True
1020                  End If
1030                  If lbContinua = True Then
                          '
1040                      lcEsNuevaHistoria = ""
                          
1050                      If idMotivoMovimiento <> 9 Then
1060                         lcEsNuevaHistoria = mo_AdminArchivoClinico.HistoriaClinicaEsNueva(rsSolicitudes!NroHistoriaClinica, rsSolicitudes!idTipoHistoria, oConexion)
1070                      End If
                          
1080                      mrs_HistoriasPorMover.AddNew
1090                      If lnSolicitudes = 1 Then
1100                         mrs_HistoriasPorMover!seleccionar = lbSeDaraSalida
1110                      Else
1120                         mrs_HistoriasPorMover!seleccionar = 0    'lbSeDaraSalida    debb-09/09/2015
1130                      End If
1140                      mrs_HistoriasPorMover!IdHistoriaSolicitada = rsSolicitudes!IdHistoriaSolicitada
1150                      mrs_HistoriasPorMover!idPaciente = rsSolicitudes!idPaciente
1160                      mrs_HistoriasPorMover!HistoriaClinica = rsSolicitudes!NroHistoriaClinica
1170                      mrs_HistoriasPorMover!Nombres = rsSolicitudes!Nombres
1180                      mrs_HistoriasPorMover!FechaSolicitud = rsSolicitudes!FechaSolicitud
1190                      mrs_HistoriasPorMover!FechaRequerida = rsSolicitudes!FechaRequerida
1200                      mrs_HistoriasPorMover!NroFolios = 0
1210                      mrs_HistoriasPorMover!idServicioDestino = rsSolicitudes!idServicioDestino
1220                      mrs_HistoriasPorMover!nombreServicioDestino = rsSolicitudes!nombreServicioDestino
1230                      mrs_HistoriasPorMover!IdServicioOrigen = lnIdOrigen
1240                      mrs_HistoriasPorMover!nombreServicioOrigen = lcOrigen
1250                      mrs_HistoriasPorMover!idTipoHistoria = rsSolicitudes!idTipoHistoria
1260                      mrs_HistoriasPorMover!IdMovimientoHistoria = rsSolicitudes!IdMovimientoHistoria
1270                      mrs_HistoriasPorMover!IdEstadoregistro = "A"
1280                      mrs_HistoriasPorMover!FormaPago = lcFormaPago
1290                      mrs_HistoriasPorMover!PagoCita = lcSql
1300                      mrs_HistoriasPorMover!idAtencion = IIf(IsNull(rsSolicitudes!idAtencion), 0, rsSolicitudes!idAtencion)
1310                      mrs_HistoriasPorMover!SeDaraSalida = lbSeDaraSalida
1320                      lcBuscaParametro.RetornaFechaServidorSQL
1330                      If idMotivoMovimiento <> 9 And Format(rsSolicitudes!FechaRequerida, sighentidades.DevuelveFechaSoloFormato_DMY) > lcBuscaParametro.RetornaFechaServidorSQL Then
1340                          ucMensajeParpadeando1.MensajeDeTexto = "La HC: " & Trim(rsSolicitudes!NroHistoriaClinica) & " tendrá CITA para el " & rsSolicitudes!FechaRequerida
1350                          ucMensajeParpadeando1.Visible = True
1360                      End If
1370                      mrs_HistoriasPorMover!nuevaHC = IIf(lcEsNuevaHistoria = "", False, True)
1380                  End If
1390                  lnIdNroHistoria = rsSolicitudes!NroHistoriaClinica
                      'mgaray201410c
1400                  If lbSeDaraSalida = False Then
1410                      Call addRowHistoriasClinicasRestringidas(oRsHistoriasSinSalida, rsSolicitudes, _
                                          lcOrigen, lcMotivoNoSalidaHistoria, lbEsServicioCostoCero, lbSeDaraSalida)
1420                  End If
1430                  lnNroCitasEnOtroServicio = 1: lcObservaciones = ""
1440                  Do While Not rsSolicitudes.EOF And lnIdNroHistoria = rsSolicitudes!NroHistoriaClinica
1450                     If lnNroCitasEnOtroServicio > 1 And idMotivoMovimiento = 1 Then
1460                           lcObservaciones = lcObservaciones & Trim(rsSolicitudes!nombreServicioDestino) & "/ "
1470                     End If
1480                     lnNroCitasEnOtroServicio = lnNroCitasEnOtroServicio + 1
1490                     rsSolicitudes.MoveNext
1500                     If rsSolicitudes.EOF Then
1510                        Exit Do
1520                     End If
1530                  Loop
1540                  If lcObservaciones <> "" Then
1550                     lcObservaciones = "Pasa a: " & lcObservaciones
1560                  End If
                      'chequea si paciente falleciò/paso a PASIVO
1570                  If lbContinua = True Then
1580                      lcFallecioPaciente = mo_ReglasFacturacion.DevuelveSiElPacienteFallecioOhistoriaPasoPasivo(lnIdPaciente1, oConexion)
1590                      If lcFallecioPaciente <> "" Then
1600                         lcObservaciones = lcFallecioPaciente
1610                      End If
1620                  End If
                      '
1630                  If lcObservaciones <> "" And lbContinua = True Then
1640                     mrs_HistoriasPorMover.Fields!Observaciones = Left(lcObservaciones, 100)
1650                  End If
1660              Loop
1670              mrs_HistoriasPorMover.Sort = "nombreServicioDestino,FechaRequerida"
                  'mgaray201410c
1680              Call darSalidaHistoriasEnServicioCostoCero(oRsHistoriasSinSalida)
1690              If MostrarHistoriasSinSalida(oRsHistoriasSinSalida) = True Then
      '            If ms_MensajeError <> "" Then
      '               MsgBox ms_MensajeError, vbInformation, Me.Caption
1700                 ms_MensajeError = ""
1710                 txtIdHistoriaClinica.Text = ""
1720                 On Error Resume Next
1730                 If lcDefaultHistoriaFicha = "F" Then
1740                     Me.txtFichaFamiliar.SetFocus
1750                 Else
1760                     txtIdHistoriaClinica.SetFocus
1770                 End If
1780              End If
1790              If (Me.OptDevolverHcXNroHistoria.Value = True Or Me.OptDevolverHcXNroHistoria.Visible = False) And _
                                                                                            ms_MensajeError <> "" Then
1800                 MsgBox ms_MensajeError, vbInformation, Me.Caption
1810              End If
1820          ElseIf rsSolicitudes.RecordCount = 0 Then
1830              Set oRsCitaPagada = mo_AdminArchivoClinico.HistoriasMovimientos(ml_idPaciente)
1840              If oRsCitaPagada.RecordCount > 0 Then
1850                 oRsCitaPagada.MoveFirst
                     '
1860                 Set oRsHCtieneMov = mo_AdminArchivoClinico.HistoriasSolicitadasSeleccionarPorIdMovimiento(oRsCitaPagada.Fields!IdMovimiento)
1870                 If oRsHCtieneMov.RecordCount > 0 Then
1880                    lcSql = "(F.Requer: " & oRsHCtieneMov.Fields!FechaRequerida & ")"
1890                 Else
1900                    lcSql = ""
1910                 End If
1920                 oRsHCtieneMov.Close
                     '
1930                 If idMotivoMovimiento = 9 Then
1940                    If oRsCitaPagada.Fields!IdMotivo <> 9 Then
1950                          ms_MensajeError = "La HC: " & Trim(txtIdHistoriaClinica.Text) & " ya SALIO el " & oRsCitaPagada!FechaMovimiento & " " & lcSql & "(" & Trim(oRsCitaPagada!ddestino) & ")"
1960                          ucMensajeParpadeando1.MensajeDeTexto = ms_MensajeError
1970                          ucMensajeParpadeando1.Visible = True
1980                    Else
1990                          ms_MensajeError = "La HC: " & Trim(txtIdHistoriaClinica.Text) & " ya RETORNO el " & oRsCitaPagada!FechaMovimiento & " " & lcSql & " (" & Trim(oRsCitaPagada!dorigen) & ")"
2000                          ucMensajeParpadeando1.MensajeDeTexto = ms_MensajeError
2010                          ucMensajeParpadeando1.Visible = True
2020                    End If
2030                 Else
2040                    If oRsCitaPagada.Fields!IdMotivo <> 9 Then
2050                          ms_MensajeError = "La HC: " & Trim(txtIdHistoriaClinica.Text) & " ya SALIO el " & oRsCitaPagada!FechaMovimiento & " " & lcSql & "(" & Trim(oRsCitaPagada!ddestino) & ")"
2060                          ucMensajeParpadeando1.MensajeDeTexto = ms_MensajeError
2070                          ucMensajeParpadeando1.Visible = True
2080                          lbContinua = False
2090                    Else
2100                          ms_MensajeError = "La HC: " & Trim(txtIdHistoriaClinica.Text) & " ya RETORNO el " & oRsCitaPagada!FechaMovimiento & " " & lcSql & "(" & Trim(oRsCitaPagada!dorigen) & ")"
2110                          ucMensajeParpadeando1.MensajeDeTexto = ms_MensajeError
2120                          ucMensajeParpadeando1.Visible = True
2130                    End If
2140                 End If
2150                 MsgBox ms_MensajeError
2160                 ms_MensajeError = ""
2170              Else
2180                 MsgBox "No existe solicitud para esta historia clinica " + Trim(txtIdHistoriaClinica.Text) & Chr(13) & "chequee la opción  FACTURACION --> ESTADO DE CUENTA", vbInformation, Me.Caption
2190              End If
2200              oRsCitaPagada.Close
2210              Exit Sub
2220          End If
2230          oConexion.Close
2240          Set oRsCitaPagada = Nothing
2250          Set oRsHCtieneMov = Nothing
2260          Set oRsPaquete = Nothing
2270          Set oConexion = Nothing
2280          Exit Sub     'debb-02/05/2016
ErrLlenarGrilla:     'debb-02/05/2016
2290     MsgBox Err.Number & " " & Err.Description & _
                sighentidades.DevuelveFuenteDeLineaDelError(Erl(), "Sub LlenarGrilladeHistoriasSeleccionadas", "movimientoHistoriaDetalle.frm")   'debb-02/05/2016
           
End Sub


Private Sub cmbIdServicio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdHistoriaClinica
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdServicio_LostFocus()
    If Me.cmdAgregaPaciente.Visible = True Then
        If mo_cmbIdServicio.BoundText <> "" Then
            btnBuscar_Click
        End If
    End If
End Sub



Private Sub cmbTurnos_LostFocus()
'    Dim lcHoraIniX As String, lcHoraFinX As String, lnFor As Integer, lbOK As Boolean
'    Dim lcNuevosTurnos As String
'    lbOK = True
'    For lnFor = 0 To 1
'       lcHoraIniX = Left(cmbTurnos.List(lnFor), 5)
'       lcHoraFinX = Mid(cmbTurnos.List(lnFor), 7, 5)
'       If Not (sighentidades.EsHora(lcHoraIniX) = True And sighentidades.EsHora(lcHoraFinX)) Then
'          lbOK = False
'       End If
'    Next
'    If lbOK = True Then
'        lcNuevosTurnos = ""
'        For lnFor = 0 To 1
'           lcHoraIniX = Left(cmbTurnos.List(lnFor), 5)
'           lcHoraFinX = Mid(cmbTurnos.List(lnFor), 7, 5)
'           If lnFor = 0 Then
'              lcNuevosTurnos = lcNuevosTurnos & lcHoraIniX & "-" & lcHoraFinX & "-Mañana/"
'           Else
'              lcNuevosTurnos = lcNuevosTurnos & lcHoraIniX & "-" & lcHoraFinX & "-Tarde"
'           End If
'        Next
'        sighentidades.TurnoMovimientoHC = lcNuevosTurnos
'    End If
ServiciosCargar
End Sub

Private Sub cmdAgregaPaciente_Click()
        Dim oDOPaciente As New doPaciente
        Dim mo_PacienteDetalle As New PacienteDetalle
        mo_PacienteDetalle.Opcion = sghAgregar
        mo_PacienteDetalle.idUsuario = ml_idUsuario
        mo_PacienteDetalle.TipoServicio = sghConsultaExterna
        mo_PacienteDetalle.lcNombrePc = mo_lcNombrePc
        mo_PacienteDetalle.lnIdTablaLISTBARITEMS = 101
        mo_PacienteDetalle.Icon = Me.Icon
        mo_PacienteDetalle.AlPulsarClicEnACEPTARdebeSalirDeVentana = True
        mo_PacienteDetalle.Show 1
        Dim oConexion As New Connection
        oConexion.Open sighentidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(mo_PacienteDetalle.idPaciente, oConexion)
        If Not oDOPaciente Is Nothing Then
            If oDOPaciente.idPaciente > 0 Then
                ml_idPaciente = oDOPaciente.idPaciente
                Me.txtIdHistoriaClinica.Text = oDOPaciente.NroHistoriaClinica
                lblApellidos.Caption = Trim(oDOPaciente.ApellidoPaterno) + " " + Trim(oDOPaciente.ApellidoMaterno) + " " + Trim(oDOPaciente.PrimerNombre)
                btnBuscar.SetFocus
            End If
        End If
        oConexion.Close
        Set oConexion = Nothing
        Unload mo_PacienteDetalle
End Sub

Private Sub Form_Initialize()
    Set mo_cmbIdMotivo.MiComboBox = cmbIdMotivo
    Set mo_cmbIdServicio.MiComboBox = cmbIdServicio
    Set mo_cmbIdServiciOrigen.MiComboBox = cmbIdServicioOrigen
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub



Private Sub grdHistoriasSeleccionadas_BeforeCellUpdate(ByVal Cell As UltraGrid.SSCell, NewValue As Variant, ByVal Cancel As UltraGrid.SSReturnBoolean)

    If mi_Opcion = sghModificar Then
        If Cell.Column.Key = "Seleccionar" Then
            If NewValue = False Then
                If Cell.Row.Cells("IdMovimientoHistoria").Value <> "" Then
                End If
            End If
        End If
    End If
End Sub

'debb-12/04/2016
Private Sub grdHistoriasSeleccionadas_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    grdHistoriasSeleccionadas.Bands(0).Columns("IdPaciente").Hidden = True
    grdHistoriasSeleccionadas.Bands(0).Columns("IdTipoHistoria").Hidden = True
    grdHistoriasSeleccionadas.Bands(0).Columns("IdHistoriaSolicitada").Hidden = True
    grdHistoriasSeleccionadas.Bands(0).Columns("IdServicioDestino").Hidden = True
    grdHistoriasSeleccionadas.Bands(0).Columns("IdEstadoRegistro").Hidden = True
    grdHistoriasSeleccionadas.Bands(0).Columns("SeDaraSalida").Hidden = True
    'mgaray201410c
    grdHistoriasSeleccionadas.Bands(0).Columns("idAtencion").Hidden = True
    grdHistoriasSeleccionadas.Bands(0).Columns("IdMovimientoHistoria").Hidden = True
    
    If Val(mo_cmbIdMotivo.BoundText) = 9 Then
        grdHistoriasSeleccionadas.Bands(0).Columns("FormaPago").Hidden = True
        grdHistoriasSeleccionadas.Bands(0).Columns("PagoCita").Hidden = True
        grdHistoriasSeleccionadas.Bands(0).Columns("nuevaHC").Hidden = True
    Else
        grdHistoriasSeleccionadas.Bands(0).Columns("FormaPago").Width = 1000
        grdHistoriasSeleccionadas.Bands(0).Columns("PagoCita").Width = 400  '800
        grdHistoriasSeleccionadas.Bands(0).Columns("nuevaHC").Width = 800
        grdHistoriasSeleccionadas.Bands(0).Columns("nuevaHC").Header.Caption = "NuevHC"
        grdHistoriasSeleccionadas.Bands(0).Columns("nuevaHC").Activation = ssActivationActivateNoEdit
    End If
    
    grdHistoriasSeleccionadas.Bands(0).Columns("HistoriaClinica").Header.Caption = "Historia"
    grdHistoriasSeleccionadas.Bands(0).Columns("HistoriaClinica").Width = 1000    '1000
    
    grdHistoriasSeleccionadas.Bands(0).Columns("Nombres").Header.Caption = "Nombres"
    grdHistoriasSeleccionadas.Bands(0).Columns("Nombres").Width = 2400
    
    grdHistoriasSeleccionadas.Bands(0).Columns("Seleccionar").Width = 500
    grdHistoriasSeleccionadas.Bands(0).Columns("IdServicioOrigen").Hidden = True
    
    grdHistoriasSeleccionadas.Bands(0).Columns("NombreServicioOrigen").Header.Caption = "Serv.Origen"
    grdHistoriasSeleccionadas.Bands(0).Columns("NombreServicioOrigen").Width = 1200  '1500
    
    grdHistoriasSeleccionadas.Bands(0).Columns("NombreServicioDestino").Header.Caption = "Servicio Destino"
    grdHistoriasSeleccionadas.Bands(0).Columns("NombreServicioDestino").Width = 3000
    
    grdHistoriasSeleccionadas.Bands(0).Columns("IdMovimientoHistoria").Header.Caption = "IdMovimiento"
    grdHistoriasSeleccionadas.Bands(0).Columns("IdMovimientoHistoria").Width = 1500

    grdHistoriasSeleccionadas.Bands(0).Columns("NroFolios").Header.Caption = "N°Folios"
    grdHistoriasSeleccionadas.Bands(0).Columns("NroFolios").Width = 500
    
    grdHistoriasSeleccionadas.Bands(0).Columns("FechaSolicitud").Header.Caption = "F.Solicitud"
    grdHistoriasSeleccionadas.Bands(0).Columns("FechaSolicitud").Width = 1000
    
    grdHistoriasSeleccionadas.Bands(0).Columns("observaciones").Width = 3500
    grdHistoriasSeleccionadas.Bands(0).Columns("FechaRequerida").Width = 1600
    
    mo_Apariencia.ConfigurarFilasBiColores grdHistoriasSeleccionadas, sighentidades.GrillaConFilasBicolor
    
End Sub



Private Sub grdHistoriasSeleccionadas_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)

     If KeyCode = vbKeyF2 Then
       Dim lnKeyCode As Integer
       lnKeyCode = KeyCode
       AdministrarKeyPreview lnKeyCode
     End If
End Sub

Private Sub OptDevolverHcXNroHistoria_Click()
    mostrarbusquedaparadevolucion
End Sub

Private Sub OptDevolverporServicio_Click()
    mostrarbusquedaparadevolucion
End Sub

Private Sub txtFechaDesde_LostFocus()
If Not esfecha(txtFechaDesde.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFechaDesde.Text = sighentidades.FECHA_VACIA_DMY
        Exit Sub
    End If
    ServiciosCargar
End Sub

Private Sub txtfechaHasta_LostFocus()
If Not esfecha(txtFechaHasta.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFechaHasta.Text = sighentidades.FECHA_VACIA_DMY
        Exit Sub
    End If
End Sub

Private Sub txtFichaFamiliar_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFichaFamiliar
    AdministrarKeyPreview KeyCode
End Sub



Private Sub txtFichaFamiliar_LostFocus()
        Dim oRsBuscar As New ADODB.Recordset
        lblApellidos.Caption = ""
        If txtFichaFamiliar.Text <> "" Then
             Set oRsBuscar = mo_AdminAdmision.PacientesSeleccionarPorFichaFamiliar(txtFichaFamiliar.Text)
             If oRsBuscar.RecordCount > 0 Then
                lblApellidos.Caption = Trim(oRsBuscar.Fields!ApellidoPaterno) + " " + Trim(oRsBuscar.Fields!ApellidoMaterno) + " " + Trim(oRsBuscar.Fields!PrimerNombre)
                ml_idPaciente = oRsBuscar.Fields!idPaciente
                Me.txtIdHistoriaClinica.Text = oRsBuscar.Fields!NroHistoriaClinica
                
             End If
             oRsBuscar.Close
             txtIdHistoriaClinica_LostFocus
        End If
        Set oRsBuscar = Nothing
End Sub

Private Sub txtIdHistoriaClinica_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdHistoriaClinica
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdHistoriaClinica_LostFocus()
    If Len(txtIdHistoriaClinica.Text) > 9 Then
       MsgBox "El Nro Historia no puede exeder de 9 números", vbInformation, "Historia"
       txtIdHistoriaClinica.Text = ""
       Exit Sub
    End If
    Dim oRsBuscar As New ADODB.Recordset
    lblApellidos.Caption = ""
    If mo_Teclado.TextoEsSoloNumeros(txtIdHistoriaClinica.Text) Then
         
         Set oRsBuscar = mo_AdminAdmision.PacientesSeleccionarPorNroHistoria(Val(txtIdHistoriaClinica.Text))
         If oRsBuscar.RecordCount > 0 Then
            lblApellidos.Caption = Trim(oRsBuscar.Fields!ApellidoPaterno) + " " + Trim(oRsBuscar.Fields!ApellidoMaterno) + " " + Trim(oRsBuscar.Fields!PrimerNombre)
            ml_idPaciente = oRsBuscar.Fields!idPaciente
         End If
         oRsBuscar.Close
    End If
    Set oRsBuscar = Nothing
    CompletarDatosDeServicioEnElLostFocus txtIdServicioDestino, Me.txtNombreServicioDestino
    mo_Formulario.MarcarComoVacio txtIdHistoriaClinica
    If lblApellidos.Caption <> "" Then
       If Me.cmdAgregaPaciente.Visible = False Then
          btnBuscar_Click
       Else
          On Error Resume Next
          Me.cmbIdServicio.SetFocus
       End If
    End If
End Sub

Private Sub txtIdHistoriaClinica_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtIdServicioDestino_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdServicioDestino
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdServicioDestino_LostFocus()
    CompletarDatosDeServicioEnElLostFocus txtIdServicioDestino, Me.txtNombreServicioDestino
    mo_Formulario.MarcarComoVacio txtIdServicioDestino
End Sub

Private Sub txtIdServicioDestino_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtIdServicioDestinoFiltro_LostFocus()
    CompletarDatosDeServicioEnElLostFocus txtIdServicioDestino, Me.txtNombreServicioDestino
    mo_Formulario.MarcarComoVacio txtIdServicioDestino
End Sub

Private Sub txtObservacion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtObservacion
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtObservacion_LostFocus()
   mo_Formulario.MarcarComoVacio txtObservacion
End Sub

Private Sub txtObservacion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdMotivo_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdMotivo
   AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdMotivo_LostFocus()
   If cmbIdMotivo.Text <> "" Then
       mo_cmbIdMotivo.BoundText = Val(Split(cmbIdMotivo.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdMotivo
End Sub

Private Sub cmbIdMotivo_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtHoraMovimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtHoraMovimiento
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtHoraMovimiento_LostFocus()
If Not sighentidades.ValidaHora(txtHoraMovimiento.Text) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
             txtHoraMovimiento.Text = sighentidades.HORA_VACIA_HM
        End If
End Sub

Private Sub txtHoraMovimiento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtFechaMovimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFechaMovimiento
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtFechaMovimiento_LostFocus()
If Not esfecha(txtFechaMovimiento.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFechaMovimiento.Text = sighentidades.FECHA_VACIA_DMY
        Exit Sub
    End If
End Sub

Private Sub txtFechaMovimiento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtIdEmpleadoRecepcion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtIdEmpleadoRecepcion
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdEmpleadoRecepcion_LostFocus()
    CompletarDatosDeEmpleadoEnElLostFocus txtIdEmpleadoRecepcion, Me.txtNombreEmpleadoRecepcion
    mo_Formulario.MarcarComoVacio txtIdEmpleadoRecepcion
End Sub

Private Sub txtIdEmpleadoRecepcion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtIdEmpleadoTransporte_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtIdEmpleadoTransporte
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdEmpleadoTransporte_LostFocus()
    CompletarDatosDeEmpleadoEnElLostFocus txtIdEmpleadoTransporte, Me.txtNombreEmpleadoTransporte
    mo_Formulario.MarcarComoVacio txtIdEmpleadoTransporte
End Sub

Private Sub txtIdEmpleadoTransporte_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtIdEmpleadoArchivo_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtIdEmpleadoArchivo
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdEmpleadoArchivo_LostFocus()
    CompletarDatosDeEmpleadoEnElLostFocus txtIdEmpleadoArchivo, Me.txtNombreEmpleadoArchivo
    mo_Formulario.MarcarComoVacio txtIdEmpleadoArchivo
    
    If Trim(Me.txtIdEmpleadoRecepcion) = "" Then Me.txtIdEmpleadoRecepcion = txtIdEmpleadoArchivo
    If Trim(Me.txtIdEmpleadoTransporte) = "" Then Me.txtIdEmpleadoTransporte = txtIdEmpleadoArchivo
    
    txtIdEmpleadoRecepcion_LostFocus
    txtIdEmpleadoTransporte_LostFocus
    
End Sub

Private Sub txtIdEmpleadoArchivo_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla MovimientosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()

    Dim oDOEmpleado As New dOEmpleado
    Set oDOEmpleado = mo_AdminComun.EmpleadosSeleccionarPorId(ml_idUsuario)
    lblArchivero = "Archivero:" + oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
    txtIdEmpleadoArchivo.Tag = oDOEmpleado.IdEmpleado
    txtIdEmpleadoArchivo.Text = oDOEmpleado.CodigoPlanilla
    txtNombreEmpleadoArchivo = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
    '
    txtIdEmpleadoRecepcion.Tag = oDOEmpleado.IdEmpleado
    txtIdEmpleadoRecepcion.Text = oDOEmpleado.CodigoPlanilla
    txtNombreEmpleadoRecepcion = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
    '
    txtIdEmpleadoTransporte.Tag = oDOEmpleado.IdEmpleado
    txtIdEmpleadoTransporte.Text = oDOEmpleado.CodigoPlanilla
    txtNombreEmpleadoTransporte = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
    '
    If lcBuscaParametro.SeleccionaFilaParametro(279) = "S" Then
       lnUsuarioFiltroCombo = ml_idUsuario
    Else
       lnUsuarioFiltroCombo = IIf(lcBuscaParametro.SeleccionaFilaParametro(231) = "S", 0, ml_idUsuario)
    End If
    mo_cmbIdMotivo.BoundText = "1"
    '
    '
    mo_Formulario.HabilitarDeshabilitar Me.txtNombreEmpleadoArchivo, False
    mo_Formulario.HabilitarDeshabilitar Me.txtNombreEmpleadoRecepcion, False
    mo_Formulario.HabilitarDeshabilitar Me.txtNombreEmpleadoTransporte, False
    mo_Formulario.HabilitarDeshabilitar Me.txtNombreServicioDestino, False

    Select Case mi_Opcion
        Case sghAgregar
        Case sghModificar
            CargarDatosALosControles2
        Case sghConsultar
            CargarDatosALosControles2
        Case sghEliminar
            CargarDatosALosControles2
    End Select
    
    Select Case mi_Opcion
        Case sghAgregar
            Me.txtFechaMovimiento.Text = Format(Date, sighentidades.DevuelveFechaSoloFormato_DMY)
            Me.txtHoraMovimiento = Format(Now, sighentidades.DevuelveHoraSoloFormato_HM)
            Me.btnListarMovimientosAsoc.Visible = False
            Me.progressRpt.Visible = False
        Case sghModificar
            Me.fraPaciente.Enabled = False
            Me.txtFechaMovimiento.Enabled = False
            Me.txtHoraMovimiento.Enabled = False
        Case sghConsultar
            Me.fraPaciente.Enabled = False
            Me.fraMovimiento.Enabled = False
            Me.btnAceptar.Enabled = False
        Case sghEliminar
            Me.fraPaciente.Enabled = False
            Me.fraMovimiento.Enabled = False
    End Select
    
    
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla MovimientosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Sub ConfiguraTurnoArchivo()
    cmbTurnos.AddItem "07:00-12:59-Mañana"
    cmbTurnos.AddItem "13:00-23:59-Tarde"
    cmbTurnos.Text = ""
    lblApellidos.Caption = ""
       
       
'    Dim lcNuevosTurnos As String, lcTurno As String
'    lcNuevosTurnos = sighentidades.TurnoMovimientoHC
'    If Len(lcNuevosTurnos) > 20 Then
'       cmbTurnos.Clear
'       lcTurno = Left(lcNuevosTurnos, InStr(lcNuevosTurnos, "/") - 1)
'       cmbTurnos.AddItem lcTurno
'       lcTurno = Mid(lcNuevosTurnos, InStr(lcNuevosTurnos, "/") + 1, 100)
'       cmbTurnos.AddItem lcTurno
'       cmbTurnos.Text = ""
'    End If
    
End Sub
Sub Form_Load()
       chkServiciosTodos.Visible = False   'debb-12/04/2016
       If lcBuscaParametro.SeleccionaFilaParametro(282) = "S" Then
          Me.lblFichaFamiliar.Visible = True
          Me.txtFichaFamiliar.Visible = True
       End If
       lcDefaultHistoriaFicha = IIf(lcBuscaParametro.SeleccionaFilaParametro(295) = "F", "F", "H")
       
       GenerarRecordsetTemporal
       cmbFecha.ListIndex = 1
       
       ConfiguraTurnoArchivo

        
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar movimiento de historia clínica"
       Case sghModificar
           Me.Caption = "Modificar movimiento de historia clínica"
       Case sghConsultar
           Me.Caption = "Consultar movimiento de historia clínica"
       Case sghEliminar
           Me.Caption = "Eliminar movimiento de historia clínica"
       End Select

       CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
       If lcBuscaParametro.SeleccionaFilaParametro(279) = "S" Then
          mo_cmbIdMotivo.BoundText = "6"
          cmdAgregaPaciente.Visible = True
       End If
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla MovimientosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Activate()
   If mi_Opcion <> sghAgregar Then
       If Not mb_ExistenDatos Then
           Me.Visible = False
           LimpiarVariablesDeMemoria
       End If
      
   End If
End Sub
Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
       Case vbKeyF1
           btnBuscarPaciente_Click
       Case vbKeyF2
           btnAceptar_Click
       Case vbKeyF6
           btnBuscar_Click
       End Select
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   AdministrarKeyPreview KeyCode
End Sub

Private Sub btnAceptar_Click()
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If AgregarDatos() Then
                   MsgBox " Los datos se agregaron exitosamente", vbInformation, Me.Caption
                   LimpiarFormulario
                   On Error Resume Next
                   If lcDefaultHistoriaFicha = "F" Then
                      Me.txtFichaFamiliar.SetFocus
                   Else
                      txtIdHistoriaClinica.SetFocus
                   End If
               Else
                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminArchivoClinico.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If ModificarDatos() Then
                   MsgBox " Los datos se modificaron exitosamente", vbInformation, Me.Caption
                   Me.Visible = False
                   LimpiarVariablesDeMemoria
               Else
                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminArchivoClinico.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                   MsgBox " Los datos se eliminaron exitosamente", vbInformation, Me.Caption
                   Me.Visible = False
                   LimpiarVariablesDeMemoria
               Else
                   MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_AdminArchivoClinico.MensajeError, vbExclamation, Me.Caption
               End If
           End If
   End Select
End Sub

Private Sub btnCancelar_Click()
   Me.Visible = False
   LimpiarVariablesDeMemoria
End Sub

Function ValidarDatosObligatorios() As Boolean
   Dim sMensaje As String
   Dim lbExisteH As Boolean
   ValidarDatosObligatorios = False
   
   If mo_cmbIdMotivo.BoundText = "" Then
       sMensaje = sMensaje + "Ingrese el motivo" + Chr(13)
   End If
   If Me.txtHoraMovimiento.Text = sighentidades.HORA_VACIA_HM Then
       sMensaje = sMensaje + "Ingrese la hora de movimiento" + Chr(13)
   End If
   If Me.txtFechaMovimiento.Text = sighentidades.FECHA_VACIA_DMY Then
       sMensaje = sMensaje + "Ingrese la fecha de movimiento" + Chr(13)
   End If
   If Me.txtIdEmpleadoRecepcion.Text = "" Then
       sMensaje = sMensaje + "Ingrese el código de planilla del empleado de recepcion" + Chr(13)
   End If
   If Me.txtIdEmpleadoTransporte.Text = "" Then
       sMensaje = sMensaje + "Ingrese el código de planilla del empleado de transporte" + Chr(13)
   End If
   If Me.txtIdEmpleadoArchivo.Text = "" Then
       sMensaje = sMensaje + "Ingrese el código de planilla del empleado de archivo" + Chr(13)
   End If
   '
   If mrs_HistoriasPorMover.RecordCount = 0 Then
       sMensaje = sMensaje + "No hay Historias para Seleccionar" + Chr(13)
   Else
        lbExisteH = False
        mrs_HistoriasPorMover.MoveFirst
        Do While Not mrs_HistoriasPorMover.EOF
             If mrs_HistoriasPorMover!seleccionar Then
                lbExisteH = True
             End If
             mrs_HistoriasPorMover.MoveNext
        Loop
        If lbExisteH = False Then
           sMensaje = sMensaje + "Seleccione al menos una Historia" + Chr(13)
        End If
   End If
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
Dim sMensaje As String

   ValidarReglas = False
   
   sMensaje = ""
   If mrs_HistoriasPorMover.RecordCount > 0 Then
     mrs_HistoriasPorMover.MoveFirst
     Do While Not mrs_HistoriasPorMover.EOF
          Select Case Val(mrs_HistoriasPorMover!idTipoHistoria)
          Case 3
              If Val(mrs_HistoriasPorMover!NroFolios) = 0 Then
                  sMensaje = sMensaje + "La historia clínica : " & mrs_HistoriasPorMover!HistoriaClinica & " es ESPECIAL." + Chr(13)
              End If
          Case 4
              If Val(mrs_HistoriasPorMover!NroFolios) = 0 Then
                  sMensaje = sMensaje + "La historia clínica : " & mrs_HistoriasPorMover!HistoriaClinica & " es JUDICIAL." + Chr(13)
              End If
          End Select
          mrs_HistoriasPorMover.MoveNext
    Loop
    If sMensaje <> "" Then
        MsgBox sMensaje + Chr(13) + "Por favor ingresar el nro de folios correspondientes", vbInformation, Me.Caption
        Exit Function
    End If
    ValidarReglas = True
  End If

End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla MovimientosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()
    
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LOS MOVIMIENTOS
    '---------------------------------------------------------------------------------
    If Not (mrs_HistoriasPorMover.BOF And mrs_HistoriasPorMover.EOF) Then
            Set mo_Movimiento = New DOMovimientoHistoriaClinica
            With mo_Movimiento
                .Observacion = Me.txtObservacion.Text
                .IdMotivo = mo_cmbIdMotivo.BoundText
                .FechaMovimiento = Format(Me.txtFechaMovimiento.Text + " " + Me.txtHoraMovimiento.Text, sighentidades.DevuelveFechaSoloFormato_DMY_HM)
                .IdEmpleadoRecepcion = Val(Me.txtIdEmpleadoRecepcion.Tag)
                .IdEmpleadoTransporte = Val(Me.txtIdEmpleadoTransporte.Tag)
                .IdEmpleadoArchivo = Val(Me.txtIdEmpleadoArchivo.Tag)
                .IdGrupoMovimiento = ml_IdGrupoMovimiento
                .IdUsuarioAuditoria = Val(Me.txtIdEmpleadoArchivo.Tag)
                .IdMovimiento = Me.IdMovimiento
            End With
    End If
    lcHCyPaciente = ""
    ms_MensajeError = Trim(Str(mrs_HistoriasPorMover.RecordCount))
    If Val(ms_MensajeError) > 0 Then
       mrs_HistoriasPorMover.MoveFirst
       If Val(ms_MensajeError) = 1 Then
          If mrs_HistoriasPorMover.Fields!seleccionar = True Then
             lcHCyPaciente = Trim(Str(mrs_HistoriasPorMover.Fields!HistoriaClinica)) & " " & Trim(mrs_HistoriasPorMover.Fields!Nombres) & " (destino: " & Trim(mrs_HistoriasPorMover.Fields!nombreServicioDestino) & ")"
          End If
       Else
          mrs_HistoriasPorMover.Find "seleccionar=true"
       End If
       If mrs_HistoriasPorMover.EOF Then
          lcHCyPaciente = Trim(Str(mrs_HistoriasPorMover.Fields!HistoriaClinica)) & " " & Trim(mrs_HistoriasPorMover.Fields!Nombres) & " (destino: " & Trim(mrs_HistoriasPorMover.Fields!nombreServicioDestino) & ")"
       End If
    End If
    ms_MensajeError = ""
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
   
   CargaDatosAlObjetosDeDatos
   If mo_Movimiento.IdMotivo <> 9 Then
      mrs_HistoriasPorMover.Filter = "SeDaraSalida=true"
   End If
   If mrs_HistoriasPorMover.RecordCount > 0 Then
      AgregarDatos = mo_AdminArchivoClinico.MovimientosHistoriaClinicaAgregar(mo_Movimiento, mrs_HistoriasPorMover, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcHCyPaciente)
   End If
   mrs_HistoriasPorMover.Filter = ""
End Function

Function DevolverHistoria_() As Boolean
   
   
   Dim oMovimiento As DOMovimientoHistoriaClinica
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LOS MOVIMIENTOS
    '---------------------------------------------------------------------------------
    Set mo_Movimientos = New Collection
    Set oMovimiento = New DOMovimientoHistoriaClinica
    With oMovimiento
        .IdMovimiento = 0
        .NroFolios = 0
        .idServicioDestino = mo_AdminComun.ParametrosIdServicioArchivoClinico()
        .IdServicioOrigen = Val(Me.txtIdServicioDestino.Tag)
        .Observacion = "Devolución al archivo"
        .IdMotivo = 9
        .FechaMovimiento = Format(Now, sighentidades.DevuelveFechaSoloFormato_DMY_HM)
        .idPaciente = mrs_HistoriasPorMover!idPaciente
        .IdEmpleadoRecepcion = ml_idUsuario
        .IdEmpleadoTransporte = 0
        .IdEmpleadoArchivo = ml_idUsuario
        .IdHistoriaSolicitada = IIf(IsNull(mrs_HistoriasPorMover!IdHistoriaSolicitada), 0, mrs_HistoriasPorMover!IdHistoriaSolicitada)
        .IdGrupoMovimiento = 0
    End With
    mo_Movimientos.Add oMovimiento
    DevolverHistoria_ = mo_AdminArchivoClinico.MovimientosHistoriaClinicaAgregar(mo_Movimiento, mrs_HistoriasPorMover, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcHCyPaciente)


End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   ModificarDatos = mo_AdminArchivoClinico.MovimientosHistoriaClinicaModificar(mo_Movimiento, mrs_HistoriasPorMover, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcHCyPaciente)

End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   EliminarDatos = mo_AdminArchivoClinico.MovimientosHistoriaClinicaEliminar(mo_Movimiento, mrs_HistoriasPorMover, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcHCyPaciente)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla MovimientosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlosControles()
Dim oDoServicio As New doServicio
Dim oConexion As New Connection
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
        Set mo_MovimientosHistoriaClinica = mo_AdminArchivoClinico.MovimientosHistoriaClinicaSeleccionarPorId(Me.IdMovimiento)
        If mo_AdminArchivoClinico.MensajeError <> "" Then
             MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminArchivoClinico.MensajeError, vbInformation, Me.Caption"
             mb_ExistenDatos = False
             Exit Sub
        End If

       If Not mo_MovimientosHistoriaClinica Is Nothing Then
           With mo_MovimientosHistoriaClinica
           
                Me.IdMovimiento = .IdMovimiento
                ml_IdGrupoMovimiento = .IdGrupoMovimiento
                Me.txtObservacion.Text = .Observacion
                mo_cmbIdMotivo.BoundText = .IdMotivo
                Me.txtHoraMovimiento.Text = Format(.FechaMovimiento, sighentidades.DevuelveHoraSoloFormato_HM)
                Me.txtFechaMovimiento.Text = Format(.FechaMovimiento, sighentidades.DevuelveFechaSoloFormato_DMY)
                
                 'Datos del paciente
                 Dim oDOPaciente As New doPaciente
                 Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(.idPaciente, oConexion)
                 If Not oDOPaciente Is Nothing Then
                     mrs_HistoriasPorMover.AddNew
                     mrs_HistoriasPorMover.Fields!idPaciente = oDOPaciente.idPaciente
                     mrs_HistoriasPorMover.Fields!HistoriaClinica = oDOPaciente.NroHistoriaClinica
                     mrs_HistoriasPorMover.Fields!Nombres = oDOPaciente.ApellidoPaterno + " " + oDOPaciente.ApellidoMaterno + " " + oDOPaciente.PrimerNombre + " " + oDOPaciente.SegundoNombre
                     mrs_HistoriasPorMover.Fields!IdServicioOrigen = .IdServicioOrigen
                     
                    Set oDoServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(.IdServicioOrigen, oConexion)
                    If Not oDoServicio Is Nothing Then
                        mrs_HistoriasPorMover.Fields!nombreServicioOrigen = oDoServicio.Codigo + " " + oDoServicio.nombre
                    Else
                        mrs_HistoriasPorMover.Fields!nombreServicioOrigen = ""
                    End If
                     
                     mrs_HistoriasPorMover.Fields!NroFolios = .NroFolios
                     
                    Dim rsHistoriaSolicitada As New Recordset
                    Set rsHistoriaSolicitada = mo_AdminArchivoClinico.HistoriasSolicitadasSeleccionarPorIdMovimiento(.IdMovimiento)
                    If Not (rsHistoriaSolicitada.EOF And rsHistoriaSolicitada.BOF) Then
                        mrs_HistoriasPorMover.Fields!IdHistoriaSolicitada = rsHistoriaSolicitada!IdHistoriaSolicitada
                        mrs_HistoriasPorMover.Fields!FechaSolicitud = rsHistoriaSolicitada!FechaSolicitud
                    End If
                    rsHistoriaSolicitada.Close
                     
                    Dim doHistoriaClinicas As New DOHistoriaClinica
                    Set doHistoriaClinicas = mo_AdminArchivoClinico.HistoriaClinicaSeleccionarPorId(oDOPaciente.NroHistoriaClinica)
                    If Not doHistoriaClinicas Is Nothing Then
                        mrs_HistoriasPorMover!idTipoHistoria = doHistoriaClinicas.idTipoHistoria
                    End If
                 
                 End If
                                
                 'Datos del servicio destino
                 Set oDoServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(.idServicioDestino, oConexion)
                 If Not oDoServicio Is Nothing Then
                     Me.txtIdServicioDestino.Tag = oDoServicio.IdServicio
                     Me.txtIdServicioDestino.Text = oDoServicio.Codigo
                     Me.txtNombreServicioDestino = oDoServicio.nombre
                 Else
                     Me.txtIdServicioDestino.Tag = ""
                     Me.txtIdServicioDestino.Text = ""
                     Me.txtNombreServicioDestino = ""
                 End If
                
                Dim oDOEmpleado As New dOEmpleado
                
                Set oDOEmpleado = mo_AdminComun.EmpleadosSeleccionarPorId(.IdEmpleadoRecepcion)
                If Not oDOEmpleado Is Nothing Then
                    Me.txtIdEmpleadoRecepcion.Tag = oDOEmpleado.IdEmpleado
                    txtIdEmpleadoRecepcion.Text = oDOEmpleado.CodigoPlanilla
                    txtNombreEmpleadoRecepcion = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
                End If
                
                Set oDOEmpleado = mo_AdminComun.EmpleadosSeleccionarPorId(.IdEmpleadoTransporte)
                If Not oDOEmpleado Is Nothing Then
                    Me.txtIdEmpleadoTransporte.Tag = oDOEmpleado.IdEmpleado
                    txtIdEmpleadoTransporte.Text = oDOEmpleado.CodigoPlanilla
                    Me.txtNombreEmpleadoTransporte = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
                End If
                
                Set oDOEmpleado = mo_AdminComun.EmpleadosSeleccionarPorId(.IdEmpleadoArchivo)
                If Not oDOEmpleado Is Nothing Then
                    Me.txtIdEmpleadoArchivo.Tag = oDOEmpleado.IdEmpleado
                    txtIdEmpleadoArchivo.Text = oDOEmpleado.CodigoPlanilla
                    Me.txtNombreEmpleadoArchivo = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
                End If
                 
                If Not mo_AdminArchivoClinico.MovimientosHistoriaEsUltimoMovimiento(oDOPaciente.idPaciente, Me.IdMovimiento) Then
                    Select Case mi_Opcion
                    Case sghModificar
                        MsgBox "No podrá modificar el servicio destino, esto sólo esta permitido si es el último movimiento", vbInformation, Me.Caption
                        'Me.btnDevolverHistoria.Enabled = False
                        Me.txtIdServicioDestino.Enabled = False
                    Case sghEliminar
                        MsgBox "Solo se puede eliminar el último elemento", vbInformation, Me.Caption
                        mi_Opcion = sghConsultar
                    End Select
                End If
                 
                 mb_ExistenDatos = True
           End With
       Else
           mb_ExistenDatos = False
           Exit Sub
       End If
   
End Sub

Sub CargarDatosALosControles2()
Dim oDoServicio As New doServicio

    Set mo_MovimientosHistoriaClinica = mo_AdminArchivoClinico.MovimientosHistoriaClinicaSeleccionarPorId(Me.IdMovimiento)
    If mo_AdminArchivoClinico.MensajeError <> "" Then
         MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminArchivoClinico.MensajeError, vbInformation, Me.Caption"
         mb_ExistenDatos = False
         Exit Sub
    End If
    
    If Not mo_MovimientosHistoriaClinica Is Nothing Then
             ml_IdGrupoMovimiento = mo_MovimientosHistoriaClinica.IdGrupoMovimiento
             Me.txtObservacion.Text = mo_MovimientosHistoriaClinica.Observacion
             mo_cmbIdMotivo.BoundText = mo_MovimientosHistoriaClinica.IdMotivo
             Me.txtHoraMovimiento.Text = Format(mo_MovimientosHistoriaClinica.FechaMovimiento, sighentidades.DevuelveHoraSoloFormato_HM)
             Me.txtFechaMovimiento.Text = Format(mo_MovimientosHistoriaClinica.FechaMovimiento, sighentidades.DevuelveFechaSoloFormato_DMY)
    
            Dim oDOEmpleado As New dOEmpleado
            
            Set oDOEmpleado = mo_AdminComun.EmpleadosSeleccionarPorId(mo_MovimientosHistoriaClinica.IdEmpleadoRecepcion)
            If Not oDOEmpleado Is Nothing Then
                Me.txtIdEmpleadoRecepcion.Tag = oDOEmpleado.IdEmpleado
                txtIdEmpleadoRecepcion.Text = oDOEmpleado.CodigoPlanilla
                txtNombreEmpleadoRecepcion = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
            End If
            
            Set oDOEmpleado = mo_AdminComun.EmpleadosSeleccionarPorId(mo_MovimientosHistoriaClinica.IdEmpleadoTransporte)
            If Not oDOEmpleado Is Nothing Then
                Me.txtIdEmpleadoTransporte.Tag = oDOEmpleado.IdEmpleado
                txtIdEmpleadoTransporte.Text = oDOEmpleado.CodigoPlanilla
                Me.txtNombreEmpleadoTransporte = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
            End If
            
            Set oDOEmpleado = mo_AdminComun.EmpleadosSeleccionarPorId(mo_MovimientosHistoriaClinica.IdEmpleadoArchivo)
            If Not oDOEmpleado Is Nothing Then
                Me.txtIdEmpleadoArchivo.Tag = oDOEmpleado.IdEmpleado
                txtIdEmpleadoArchivo.Text = oDOEmpleado.CodigoPlanilla
                Me.txtNombreEmpleadoArchivo = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
            End If
    
            mb_ExistenDatos = True
    Else
        mb_ExistenDatos = False
    End If
    
    'Detalle del movimiento
    Dim oRsCitaPagada As New ADODB.Recordset
    Dim lcSql As String
    Dim lcFormaPago As String
    Dim rsMovimientoDeHistorias As New Recordset
    Set rsMovimientoDeHistorias = mo_AdminArchivoClinico.MovimientosHistoriasClinicasPorIdGrupo(mo_MovimientosHistoriaClinica.IdGrupoMovimiento)
    If mo_AdminArchivoClinico.MensajeError <> "" Then
         MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminArchivoClinico.MensajeError, vbInformation, Me.Caption"
         mb_ExistenDatos = False
         Exit Sub
    End If
    If mi_Opcion = sghEliminar Then
       rsMovimientoDeHistorias.Filter = "idPaciente=" & ml_idPacienteSeleccionado
    End If
    Do While Not rsMovimientoDeHistorias.EOF
            lcSql = " ": lcFormaPago = ""
            If Not IsNull(rsMovimientoDeHistorias.Fields!idAtencion) Then
                Set oRsCitaPagada = mo_AdminArchivoClinico.HistoriaPagoCita(rsMovimientoDeHistorias.Fields!idAtencion, "1")
                 lcSql = " "
                 If oRsCitaPagada.RecordCount > 0 Then
                    If oRsCitaPagada.Fields!idestadofacturacion = 4 Then
                       lcSql = "Si"
                    End If
                    lcFormaPago = oRsCitaPagada!Descripcion
                 End If
                 oRsCitaPagada.Close
            End If
            
            mrs_HistoriasPorMover.AddNew
            
            mrs_HistoriasPorMover!seleccionar = True
            mrs_HistoriasPorMover!IdHistoriaSolicitada = rsMovimientoDeHistorias!IdHistoriaSolicitada
            mrs_HistoriasPorMover!idPaciente = rsMovimientoDeHistorias!idPaciente
            mrs_HistoriasPorMover!HistoriaClinica = rsMovimientoDeHistorias!NroHistoriaClinica
            mrs_HistoriasPorMover!Nombres = rsMovimientoDeHistorias!Nombres
            mrs_HistoriasPorMover!FechaSolicitud = rsMovimientoDeHistorias!FechaSolicitud
            mrs_HistoriasPorMover!FechaRequerida = rsMovimientoDeHistorias!FechaRequerida
            mrs_HistoriasPorMover!NroFolios = rsMovimientoDeHistorias!NroFolios
            mrs_HistoriasPorMover!idServicioDestino = IIf(IsNull(rsMovimientoDeHistorias!idServicioDestino), 0, rsMovimientoDeHistorias!idServicioDestino)
            mrs_HistoriasPorMover!nombreServicioDestino = IIf(IsNull(rsMovimientoDeHistorias!nombreServicioDestino), "", rsMovimientoDeHistorias!nombreServicioDestino)
            mrs_HistoriasPorMover!IdServicioOrigen = IIf(IsNull(rsMovimientoDeHistorias!IdServicioOrigen), 0, rsMovimientoDeHistorias!IdServicioOrigen)
            mrs_HistoriasPorMover!nombreServicioOrigen = IIf(IsNull(rsMovimientoDeHistorias!nombreServicioOrigen), "", rsMovimientoDeHistorias!nombreServicioOrigen)
            mrs_HistoriasPorMover!idTipoHistoria = rsMovimientoDeHistorias!idTipoHistoria
            mrs_HistoriasPorMover!IdMovimientoHistoria = rsMovimientoDeHistorias!IdMovimientoHistoria
            
            mrs_HistoriasPorMover!IdEstadoregistro = "M"
            mrs_HistoriasPorMover!FormaPago = lcFormaPago
            mrs_HistoriasPorMover!idAtencion = IIf(IsNull(rsMovimientoDeHistorias!idAtencion), 0, rsMovimientoDeHistorias!idAtencion)
            mrs_HistoriasPorMover!PagoCita = lcSql
            rsMovimientoDeHistorias.MoveNext
    Loop
    
   
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla MovimientosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()

            
            Me.IdMovimiento = 0
            Me.txtIdServicioDestino.Text = ""
            Me.txtNombreServicioDestino.Text = ""
            Me.txtObservacion.Text = ""
            Me.txtIdHistoriaClinica.Text = ""
            txtFichaFamiliar.Text = ""
            lblApellidos.Caption = ""
            ucMensajeParpadeando1.MensajeDeTexto = ""
            ucMensajeParpadeando1.Visible = False
            
            Me.txtFechaMovimiento.Text = Format(Date, sighentidades.DevuelveFechaSoloFormato_DMY)
            Me.txtHoraMovimiento.Text = Format(Now, sighentidades.DevuelveHoraSoloFormato_HM)
            If mrs_HistoriasPorMover.RecordCount > 0 Then
                mrs_HistoriasPorMover.MoveFirst
                Do While Not mrs_HistoriasPorMover.EOF
                    mrs_HistoriasPorMover.Delete
                    mrs_HistoriasPorMover.Update
                    mrs_HistoriasPorMover.MoveNext
                Loop
            End If
            On Error Resume Next
            If lcDefaultHistoriaFicha = "F" Then
                Me.txtFichaFamiliar.SetFocus
            Else
                Me.txtIdHistoriaClinica.SetFocus
            End If
End Sub


Sub GenerarRecordsetTemporal()
    
    With mrs_HistoriasPorMover
          .Fields.Append "Seleccionar", adBoolean
          .Fields.Append "IdHistoriaSolicitada", adInteger, 4, adFldIsNullable
          .Fields.Append "IdPaciente", adInteger
          .Fields.Append "NuevaHC", adBoolean
          .Fields.Append "HistoriaClinica", adInteger
          .Fields.Append "Nombres", adVarChar, 255
          .Fields.Append "FormaPago", adVarChar, 100, adFldIsNullable
          .Fields.Append "PagoCita", adVarChar, 30, adFldIsNullable
'          .Fields.Append "FechaSolicitud", adVarChar, 255, adFldIsNullable
          .Fields.Append "FechaRequerida", adVarChar, 255, adFldIsNullable
'          .Fields.Append "NroFolios", adInteger, 4, adFldIsNullable
          .Fields.Append "IdServicioOrigen", adInteger, , adFldIsNullable
          .Fields.Append "NombreServicioOrigen", adVarChar, 100, adFldIsNullable
          .Fields.Append "IdServicioDestino", adInteger
          .Fields.Append "NombreServicioDestino", adVarChar, 100
          .Fields.Append "IdTipoHistoria", adInteger
          .Fields.Append "IdMovimientoHistoria", adInteger, 4, adFldIsNullable
          .Fields.Append "IdEstadoregistro", adChar, 1
          .Fields.Append "idAtencion", adInteger
          .Fields.Append "SeDaraSalida", adBoolean
          .Fields.Append "Observaciones", adVarChar, 100, adFldIsNullable
          .Fields.Append "NroFolios", adInteger, 4, adFldIsNullable
          .Fields.Append "FechaSolicitud", adVarChar, 255, adFldIsNullable
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
    
    Set Me.grdHistoriasSeleccionadas.DataSource = mrs_HistoriasPorMover
    
End Sub


Private Sub btnAgregarDx_Click()
Dim oDOPaciente As New doPaciente

    Me.txtIdHistoriaClinica = Trim(Me.txtIdHistoriaClinica)
    
    If Me.txtIdHistoriaClinica = "" Then
        MsgBox "Por favor ingresar el Nro de historia clínica", vbInformation, Me.Caption
        Exit Sub
    End If
    
    If Val(mo_cmbIdMotivo.BoundText) = 0 Then
        MsgBox "Ingrese el motivo del movimiento", vbInformation, Me.Caption
        Exit Sub
    End If
    
    Select Case Val(mo_cmbIdMotivo.BoundText)
    Case 1, 2, 3, 9, 10 'Consultorios externos, Hospitalizacion, Emergencia,Devolucion archivo
    Case 4, 5, 6, 7, 8 'Investigacion, Docencia, Tramites administrativos, Interconsultas
        If Val(Me.txtIdServicioDestino.Tag) = 0 Then
            MsgBox "Por favor ingresar el servicio destino", vbInformation, Me.Caption
            Exit Sub
        End If
    End Select
    

    Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorHistoriaClinicaDefinitiva(Me.txtIdHistoriaClinica)
    If oDOPaciente.idPaciente = 0 Then
        MsgBox "No existe un paciente con el nro de historia ingresado", vbInformation, Me.Caption
        Exit Sub
    End If
    
    'Verificar si ya existe
    If mrs_HistoriasPorMover.RecordCount > 0 Then
        mrs_HistoriasPorMover.MoveFirst
        Do While Not mrs_HistoriasPorMover.EOF
            If mrs_HistoriasPorMover!HistoriaClinica = Me.txtIdHistoriaClinica Then
                MsgBox "El N° de historia clínica ingresado ya se ha seleccionado", vbInformation, Me.Caption
                Exit Sub
            End If
            mrs_HistoriasPorMover.MoveNext
        Loop
    End If
    
    With mrs_HistoriasPorMover

        Dim lIdHistoriaSolicitada As Long
        Dim daFechaSolicitud As Date
        
        'Valida que exista una solicitud
        Select Case Val(mo_cmbIdMotivo.BoundText)
        Case 1, 2, 3 'Consultorios externos, Hospitalizacion, Emergencia,Devolucion archivo
 
        Case 4, 5, 6, 7, 8, 9, 10  'Investigacion, Docencia, Tramites administrativos, Interconsultas
                    
            'Detalle del movimiento
            Dim rsMovimientoDeHistorias As New Recordset
            Set rsMovimientoDeHistorias = mo_AdminArchivoClinico.MovimientosHistoriasClinicasDetallePorIdPaciente(oDOPaciente.idPaciente)
        
            Do While Not rsMovimientoDeHistorias.EOF
                  mrs_HistoriasPorMover!IdHistoriaSolicitada = rsMovimientoDeHistorias!IdHistoriaSolicitada
                  mrs_HistoriasPorMover!idPaciente = rsMovimientoDeHistorias!idPaciente
                  mrs_HistoriasPorMover!HistoriaClinica = rsMovimientoDeHistorias!HistoriaClinica
                  mrs_HistoriasPorMover!Nombres = rsMovimientoDeHistorias!Nombres
                  mrs_HistoriasPorMover!FechaSolicitud = rsMovimientoDeHistorias!FechaSolicitud
                  mrs_HistoriasPorMover!FechaRequerida = rsMovimientoDeHistorias!FechaRequerida
                  mrs_HistoriasPorMover!NroFolios = rsMovimientoDeHistorias!NroFolios
                  mrs_HistoriasPorMover!idServicioDestino = Val(Me.txtIdServicioDestino.Tag)
                  mrs_HistoriasPorMover!nombreServicioDestino = Me.txtNombreServicioDestino
                  mrs_HistoriasPorMover!IdServicioOrigen = rsMovimientoDeHistorias!IdServicioOrigen
                  mrs_HistoriasPorMover!nombreServicioOrigen = rsMovimientoDeHistorias!nombreServicioOrigen
                  mrs_HistoriasPorMover!idTipoHistoria = rsMovimientoDeHistorias!idTipoHistoria
                  mrs_HistoriasPorMover!IdMovimientoHistoria = rsMovimientoDeHistorias!IdMovimientoHistoria
                  mrs_HistoriasPorMover!IdEstadoregistro = "A"
            Loop
        End Select

        End With
    
    Me.txtIdHistoriaClinica = ""

End Sub

Private Sub btnQuitarDx_Click()
    On Error Resume Next
    With mrs_HistoriasPorMover
        If Not .EOF And Not .BOF Then
           .Delete
           .Update
        End If
    End With
End Sub

Sub CompletarDatosResponsable(txtIdResponsable As TextBox, txtNombreResponsable As TextBox)
'Dim oBusqueda As New EmpleadosBusqueda
Dim oBusqueda As New SIGHNegocios.BuscaEmpleados
Dim oDOEmpleado As New dOEmpleado
    oBusqueda.MostrarFormulario
    'oBusqueda.Show 1
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOEmpleado = mo_AdminComun.EmpleadosSeleccionarPorId(oBusqueda.idRegistroSeleccionado)
        If Not oDOEmpleado Is Nothing Then
            txtIdResponsable.Tag = oDOEmpleado.IdEmpleado
            txtIdResponsable.Text = oDOEmpleado.CodigoPlanilla
            txtNombreResponsable = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        End If
    End If

End Sub
Sub CompletarDatosDeServicio(txtIdServicio As TextBox, lblDescripcionServicio As TextBox)
Dim oBusqueda As New SIGHNegocios.BuscaServicioHosp
Dim oDoServicio As New doServicio
Dim oConexion As New Connection
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    oBusqueda.HabilitarTipoServicio = True
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDoServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oConexion)
        If Not oDoServicio Is Nothing Then
            txtIdServicio.Text = oDoServicio.Codigo
            txtIdServicio.Tag = oDoServicio.IdServicio
            lblDescripcionServicio = oDoServicio.nombre
        Else
            txtIdServicio.Text = ""
            txtIdServicio.Tag = ""
            lblDescripcionServicio = ""
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oBusqueda = Nothing
    Set oDoServicio = Nothing

End Sub

Sub CompletarDatosDeEmpleadoEnElLostFocus(txtCodigoPlanilla As TextBox, txtNombre As TextBox)
Dim oDOEmpleado As New dOEmpleado

        If mo_AdminComun.EmpleadosSeleccionarPorCodigo(txtCodigoPlanilla.Text, oDOEmpleado) Then
            txtCodigoPlanilla.Tag = oDOEmpleado.IdEmpleado
            txtCodigoPlanilla.Text = oDOEmpleado.CodigoPlanilla
            txtNombre = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        Else
            txtCodigoPlanilla.Tag = ""
            txtCodigoPlanilla = ""
            txtNombre = ""
        End If
End Sub

Sub CompletarDatosDeServicioEnElLostFocus(txtIdServicio As TextBox, lblDescripcionServicio As TextBox)
    
    txtIdServicio.Text = UCase(txtIdServicio.Text)
    If txtIdServicio.Text <> "" Then
        Dim oDoServicio As doServicio
        Set oDoServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorCodigo(txtIdServicio.Text)
        If Not oDoServicio Is Nothing Then
            txtIdServicio.Tag = oDoServicio.IdServicio
            lblDescripcionServicio.Text = oDoServicio.nombre
        Else
            txtIdServicio.Tag = ""
            lblDescripcionServicio.Text = ""
        End If
   End If

End Sub

Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_Teclado = Nothing
    Set mo_Formulario = Nothing
    Set mo_MovimientosHistoriaClinica = Nothing
    Set mo_AdminAdmision = Nothing
    Set mo_AdminArchivoClinico = Nothing
    Set mo_AdminServiciosHosp = Nothing
    Set mo_AdminComun = Nothing
    Set mrs_HistoriasPorMover = Nothing
    Set mo_Movimientos = Nothing
    Set mo_cmbIdMotivo = Nothing
    Set mo_cmbIdServicio = Nothing
    Set mo_cmbIdServiciOrigen = Nothing
    Set mo_Apariencia = Nothing
    Set lcBuscaParametro = Nothing
End Sub

Sub mostrarbusquedaparadevolucion()
    If OptDevolverHcXNroHistoria.Value = True Then
        cmbIdServicio.ListIndex = -1
        frmFiltro2.Visible = False
        Me.cmbCondicionFechas.Enabled = True
        cmbIdServicio.Enabled = True
        cmbFecha.ListIndex = 2
        Me.Label1.Visible = True
        Me.txtIdHistoriaClinica.Visible = True
        Me.btnBuscarPaciente.Visible = True
        Me.lblApellidos.Visible = True
    Else
        txtIdHistoriaClinica.Text = ""
        frmFiltro2.Visible = True
        Me.cmbCondicionFechas.Enabled = True
        cmbIdServicio.Enabled = True
        cmbFecha.ListIndex = 2
        Label1.Visible = False
        txtIdHistoriaClinica.Visible = False
        lblFichaFamiliar.Visible = False
        txtFichaFamiliar.Visible = False
        btnBuscarPaciente.Visible = False
        lblApellidos.Visible = False
    End If
End Sub

'mgaray201410c
Private Function inicializarRsHistoriasClinicasRestringidas() As ADODB.Recordset
    Dim oRs As New ADODB.Recordset
    With oRs
        .Fields.Append "IdHistoriaSolicitada", adInteger
        .Fields.Append "HistoriaClinica", adInteger
        .Fields.Append "Nombres", adVarChar, 255
        .Fields.Append "NombreServicioOrigen", adVarChar, 200, adFldIsNullable
        .Fields.Append "NombreServicioDestino", adVarChar, 200, adFldIsNullable
        .Fields.Append "Motivo", adVarChar, 200, adFldIsNullable
        .Fields.Append "EsServicioCostoCero", adBoolean
        .Fields.Append "SeDaraSalida", adBoolean
        
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
    Set inicializarRsHistoriasClinicasRestringidas = oRs
End Function

Private Function addRowHistoriasClinicasRestringidas(ByRef oRs As ADODB.Recordset, _
                    rsSolicitudes As ADODB.Recordset, lcOrigen As String, _
                    lcMotivo As String, lbEsServicioCostoCero As Boolean, _
                    lbSeDaraSalida As Boolean)
On Error GoTo miError
    With oRs
        .AddNew
        .Fields!IdHistoriaSolicitada = rsSolicitudes!IdHistoriaSolicitada
        .Fields!HistoriaClinica = rsSolicitudes!NroHistoriaClinica
        .Fields!Nombres = rsSolicitudes!Nombres
        .Fields!nombreServicioOrigen = lcOrigen
        .Fields!nombreServicioDestino = rsSolicitudes!nombreServicioDestino
        .Fields!Motivo = lcMotivo
        .Fields!EsServicioCostoCero = lbEsServicioCostoCero
        .Fields!SeDaraSalida = lbSeDaraSalida
        .Update
    End With
miError:
    If Err Then
        MsgBox Err.Number & " : " & Err.Description, vbCritical, "Movimiento de H.C"
    End If
End Function

Private Function darSalidaHistoriasEnServicioCostoCero(oRsHistoriasSinSalida As ADODB.Recordset)
    If oRsHistoriasSinSalida.RecordCount > 0 Then
        oRsHistoriasSinSalida.Filter = "EsServicioCostoCero=1"
        If oRsHistoriasSinSalida.RecordCount > 0 Then
            'setear recorset filtrado a formulario
            Dim oFormHistorias As New MovimientoHistoriasNoListas
            Set oFormHistorias.RsHistorias = oRsHistoriasSinSalida
            oFormHistorias.Caption = "Movimiento de Historias Clinicas"
            oFormHistorias.lblTitulo.Caption = "Las siguientes Historias Se solicitan en un servicio que acepta Costo cero y no cancelaron cita. ¿ desea darle SALIDA de todas maneras ?"
            oFormHistorias.mostrarBotonesSiNO
            oFormHistorias.Show 1
            If oFormHistorias.Respuesta = True Then
                Call darSalidaAHistoriasEnServicioCostoCero(oRsHistoriasSinSalida)
            End If
            Unload oFormHistorias
        End If
    End If
End Function

Private Function darSalidaAHistoriasEnServicioCostoCero(oRsHistoriasSinSalida As ADODB.Recordset)
    oRsHistoriasSinSalida.MoveFirst
    While oRsHistoriasSinSalida.EOF = False
        mrs_HistoriasPorMover.MoveFirst
        mrs_HistoriasPorMover.Find "IdHistoriaSolicitada=" & oRsHistoriasSinSalida.Fields!IdHistoriaSolicitada
        If mrs_HistoriasPorMover.EOF = False Then
            mrs_HistoriasPorMover.Fields!seleccionar = True
            mrs_HistoriasPorMover.Fields!SeDaraSalida = True
            mrs_HistoriasPorMover.Update
        End If
        oRsHistoriasSinSalida.Fields!SeDaraSalida = True
        oRsHistoriasSinSalida.Update
        oRsHistoriasSinSalida.MoveNext
    Wend
End Function

Private Function MostrarHistoriasSinSalida(oRsHistoriasSinSalida As ADODB.Recordset) As Boolean
    MostrarHistoriasSinSalida = False
    
    If oRsHistoriasSinSalida.RecordCount > 0 Then
        oRsHistoriasSinSalida.Filter = "SeDaraSalida=0"
        If oRsHistoriasSinSalida.RecordCount > 0 Then
            MostrarHistoriasSinSalida = True
            'setear recorset filtrado a formulario
            Dim oFormHistorias As New MovimientoHistoriasNoListas
            Set oFormHistorias.RsHistorias = oRsHistoriasSinSalida
            oFormHistorias.Caption = "Movimiento de Historias Clinicas"
            oFormHistorias.lblTitulo.Caption = "Las siguentes Historias no podrá darle SALIDA"
            oFormHistorias.mostrarBotonesSoloAceptar
            oFormHistorias.Show 1
            Unload oFormHistorias
        End If
    End If
End Function
