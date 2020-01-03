VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form ComprobantePagoDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tablas"
   ClientHeight    =   9315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   13155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   9165
      Left            =   30
      TabIndex        =   0
      Top             =   120
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   16166
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tablas"
      TabPicture(0)   =   "ComprobantePagoDetalle.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Facturacion Farmacia"
      TabPicture(1)   =   "ComprobantePagoDetalle.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "grdFacturacionBienesFinanciamiento"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtNroCuenta"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command19"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command14"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Facturacion Servicios"
      TabPicture(2)   =   "ComprobantePagoDetalle.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "grdFacturacionServicioFinanciamientos"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame4"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame5"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "txtCuentaS"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Command10"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Command24"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "HBT- Actualizar Datos"
      TabPicture(3)   =   "ComprobantePagoDetalle.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblProcesando"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "ProgressBar1"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "cmbConsideraciones"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Frame6"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Frame7"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Frame8"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Frame9"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).ControlCount=   7
      Begin VB.Frame Frame9 
         Height          =   1485
         Left            =   -66960
         TabIndex        =   67
         Top             =   3960
         Width           =   4785
         Begin VB.CheckBox chkHistorias 
            Caption         =   "Actualiza los datos de Movimientos de Historias ya grabadas?"
            Height          =   345
            Left            =   120
            TabIndex        =   69
            Top             =   270
            Width           =   4575
         End
         Begin VB.CommandButton cmdHistorias 
            Caption         =   "3) Agrega los Movimiento de Historias que faltan"
            Height          =   525
            Left            =   90
            TabIndex        =   68
            Top             =   810
            Width           =   4575
         End
      End
      Begin VB.Frame Frame8 
         Height          =   1485
         Left            =   -66960
         TabIndex        =   63
         Top             =   2310
         Width           =   4785
         Begin VB.CommandButton cmdProgramacion 
            Caption         =   "2) Agrega los Programación que faltan"
            Height          =   525
            Left            =   90
            TabIndex        =   65
            Top             =   870
            Width           =   4575
         End
         Begin VB.CheckBox chkProgramacion 
            Caption         =   "Actualiza los datos de Programació ya grabados?"
            Height          =   345
            Left            =   120
            TabIndex        =   64
            Top             =   270
            Width           =   3975
         End
      End
      Begin VB.Frame Frame7 
         Height          =   2625
         Left            =   -66960
         TabIndex        =   55
         Top             =   5580
         Width           =   4785
         Begin VB.CommandButton cmdProcesaAtenciones 
            Caption         =   "4) Agrega Movimiento de Atenciones"
            Height          =   525
            Left            =   90
            TabIndex        =   62
            Top             =   1860
            Width           =   4605
         End
         Begin VB.TextBox txtOdbc 
            Height          =   345
            Left            =   1650
            TabIndex        =   57
            Text            =   "GalenhosHBT"
            Top             =   240
            Width           =   3015
         End
         Begin MSMask.MaskEdBox txtFinicial 
            Height          =   345
            Left            =   1680
            TabIndex        =   60
            Top             =   840
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFfinal 
            Height          =   345
            Left            =   3510
            TabIndex        =   61
            Top             =   840
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "hasta"
            Height          =   195
            Left            =   3030
            TabIndex        =   59
            Top             =   900
            Width           =   390
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "F.tabla Atenciones"
            Height          =   195
            Left            =   120
            TabIndex        =   58
            Top             =   900
            Width           =   1320
         End
         Begin VB.Label Label3 
            Caption         =   "ODBC del Servidor:"
            Height          =   285
            Left            =   90
            TabIndex        =   56
            Top             =   270
            Width           =   1515
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1575
         Left            =   -66960
         TabIndex        =   53
         Top             =   750
         Width           =   4785
         Begin VB.OptionButton optPacienteAdd 
            Caption         =   "Adiciona nuevos Pacientes"
            Height          =   285
            Left            =   120
            TabIndex        =   71
            Top             =   210
            Value           =   -1  'True
            Width           =   4545
         End
         Begin VB.OptionButton optPacienteAct 
            Caption         =   "Actualiza los datos de Pacientes ya grabados y Adiciona Nuevos"
            Height          =   345
            Left            =   120
            TabIndex        =   70
            Top             =   540
            Width           =   4545
         End
         Begin VB.CommandButton cmdPacientes 
            Caption         =   "1) Agrega los Pacientes que faltan"
            Height          =   525
            Left            =   90
            TabIndex        =   54
            Top             =   1020
            Width           =   4575
         End
      End
      Begin VB.ListBox cmbConsideraciones 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   7410
         Left            =   -74940
         TabIndex        =   51
         Top             =   810
         Width           =   7905
      End
      Begin VB.CommandButton Command14 
         Caption         =   "New"
         Height          =   255
         Left            =   -62640
         TabIndex        =   45
         Top             =   2760
         Width           =   465
      End
      Begin VB.CommandButton Command24 
         Caption         =   "New"
         Height          =   255
         Left            =   -62490
         TabIndex        =   40
         Top             =   2760
         Width           =   465
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Del"
         Height          =   255
         Left            =   -62640
         TabIndex        =   35
         Top             =   3120
         Width           =   375
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Del"
         Height          =   255
         Left            =   -62460
         TabIndex        =   28
         Top             =   3090
         Width           =   375
      End
      Begin VB.TextBox txtCuentaS 
         Height          =   315
         Left            =   -73800
         TabIndex        =   24
         Top             =   8640
         Width           =   1515
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1755
         Left            =   -74970
         TabIndex        =   22
         Top             =   390
         Width           =   12855
         Begin VB.CommandButton Command23 
            Caption         =   "New"
            Height          =   255
            Left            =   12360
            TabIndex        =   39
            Top             =   270
            Width           =   465
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Del"
            Height          =   255
            Left            =   12360
            TabIndex        =   27
            Top             =   600
            Width           =   375
         End
         Begin MSDataGridLib.DataGrid grdFactOrdenServicio 
            Height          =   1485
            Left            =   60
            TabIndex        =   23
            Top             =   180
            Width           =   12255
            _ExtentX        =   21616
            _ExtentY        =   2619
            _Version        =   393216
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Cabecera Despacho (FACTORDENSERVICIO)"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         ForeColor       =   &H80000008&
         Height          =   4635
         Left            =   -74970
         TabIndex        =   19
         Top             =   3960
         Width           =   12855
         Begin VB.CommandButton Command27 
            Caption         =   "New"
            Height          =   255
            Left            =   12330
            TabIndex        =   43
            Top             =   3180
            Width           =   465
         End
         Begin VB.CommandButton Command26 
            Caption         =   "New"
            Height          =   255
            Left            =   12330
            TabIndex        =   42
            Top             =   1470
            Width           =   465
         End
         Begin VB.CommandButton Command25 
            Caption         =   "New"
            Height          =   255
            Left            =   12330
            TabIndex        =   41
            Top             =   180
            Width           =   465
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Del"
            Height          =   255
            Left            =   12360
            TabIndex        =   32
            Top             =   3510
            Width           =   375
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Del"
            Height          =   255
            Left            =   12420
            TabIndex        =   30
            Top             =   1800
            Width           =   375
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Del"
            Height          =   255
            Left            =   12330
            TabIndex        =   29
            Top             =   480
            Width           =   375
         End
         Begin MSDataGridLib.DataGrid grdFactOrdenServicioPagos 
            Height          =   1245
            Left            =   60
            TabIndex        =   20
            Top             =   150
            Width           =   12195
            _ExtentX        =   21511
            _ExtentY        =   2196
            _Version        =   393216
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Cabecera Pagos (FactOrdenesServicioPagos)"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid grdFacturacionServicioPagos 
            Height          =   1665
            Left            =   60
            TabIndex        =   21
            Top             =   1410
            Width           =   12255
            _ExtentX        =   21616
            _ExtentY        =   2937
            _Version        =   393216
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Detalle Pagos (FacturacionServicioPagos)"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid grdCajaComprobantesPagoS 
            Height          =   1365
            Left            =   60
            TabIndex        =   33
            Top             =   3120
            Width           =   12255
            _ExtentX        =   21616
            _ExtentY        =   2408
            _Version        =   393216
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Recibos Servicios (CajaComprobantePago)"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         ForeColor       =   &H80000008&
         Height          =   4515
         Left            =   -74940
         TabIndex        =   16
         Top             =   4200
         Width           =   12855
         Begin VB.CommandButton Command17 
            Caption         =   "New"
            Height          =   255
            Left            =   12360
            TabIndex        =   48
            Top             =   3210
            Width           =   465
         End
         Begin VB.CommandButton Command16 
            Caption         =   "New"
            Height          =   255
            Left            =   12330
            TabIndex        =   47
            Top             =   1530
            Width           =   465
         End
         Begin VB.CommandButton Command15 
            Caption         =   "New"
            Height          =   255
            Left            =   12330
            TabIndex        =   46
            Top             =   180
            Width           =   465
         End
         Begin VB.CommandButton Command22 
            Caption         =   "Del"
            Height          =   255
            Left            =   12360
            TabIndex        =   38
            Top             =   3510
            Width           =   375
         End
         Begin VB.CommandButton Command21 
            Caption         =   "Del"
            Height          =   255
            Left            =   12420
            TabIndex        =   37
            Top             =   1830
            Width           =   375
         End
         Begin VB.CommandButton Command20 
            Caption         =   "Del"
            Height          =   255
            Left            =   12360
            TabIndex        =   36
            Top             =   480
            Width           =   375
         End
         Begin MSDataGridLib.DataGrid grdFactOrdenesBienes 
            Height          =   1245
            Left            =   60
            TabIndex        =   17
            Top             =   150
            Width           =   12225
            _ExtentX        =   21564
            _ExtentY        =   2196
            _Version        =   393216
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Cabecera Pagos (FactOrdenesBienes)"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid grdFacturacionBienesPagos 
            Height          =   1575
            Left            =   60
            TabIndex        =   18
            Top             =   1500
            Width           =   12225
            _ExtentX        =   21564
            _ExtentY        =   2778
            _Version        =   393216
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Detalle Pagos (FacturacionBienesPagos)"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid grdCajaComprobantesPago 
            Height          =   1245
            Left            =   60
            TabIndex        =   31
            Top             =   3180
            Width           =   12285
            _ExtentX        =   21669
            _ExtentY        =   2196
            _Version        =   393216
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Recibos Farmacia (CajaComprobantePago)"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1905
         Left            =   -74940
         TabIndex        =   14
         Top             =   420
         Width           =   12855
         Begin VB.CommandButton Command13 
            Caption         =   "New"
            Height          =   255
            Left            =   12330
            TabIndex        =   44
            Top             =   180
            Width           =   465
         End
         Begin VB.CommandButton Command18 
            Caption         =   "Del"
            Height          =   255
            Left            =   12330
            TabIndex        =   34
            Top             =   480
            Width           =   375
         End
         Begin MSDataGridLib.DataGrid grdFarmMovimientoVentas 
            Height          =   1635
            Left            =   60
            TabIndex        =   15
            Top             =   180
            Width           =   12225
            _ExtentX        =   21564
            _ExtentY        =   2884
            _Version        =   393216
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Cabecera Despacho (FARMMOVIMIENTOVENTAS)"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin VB.TextBox txtNroCuenta 
         Height          =   315
         Left            =   -73800
         TabIndex        =   12
         Top             =   8730
         Width           =   1515
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Tablas: GalenHos"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   7875
         Left            =   90
         TabIndex        =   1
         Top             =   720
         Width           =   12945
         Begin VB.CommandButton Command32 
            Caption         =   "..."
            Height          =   345
            Left            =   12120
            TabIndex        =   74
            ToolTipText     =   "Arregla BOLETAS ANULADAS de FArmacia (originada de una  Preventas)"
            Top             =   5910
            Width           =   405
         End
         Begin VB.CommandButton Command31 
            Caption         =   "..."
            Height          =   375
            Left            =   11460
            TabIndex        =   73
            ToolTipText     =   "busca N° Historia para Pacientes con Historia=NULL"
            Top             =   6630
            Width           =   675
         End
         Begin VB.CommandButton Command30 
            Caption         =   "..."
            Height          =   255
            Left            =   9960
            TabIndex        =   72
            ToolTipText     =   "Actualiza Datos Personales de Pacientes en tablas de Laboratorio e Imagenes"
            Top             =   6720
            Width           =   555
         End
         Begin VB.CommandButton Command29 
            Caption         =   "..."
            Height          =   225
            Left            =   6930
            TabIndex        =   50
            ToolTipText     =   "PREVENTAS: Genera Punto de Carga si no existe (incluye IdServicio) y lo asocia a cada  Procedimiento CPT"
            Top             =   6750
            Width           =   405
         End
         Begin VB.CommandButton Command28 
            Caption         =   "..."
            Height          =   195
            Left            =   5040
            TabIndex        =   49
            ToolTipText     =   "Actualiza 'Convenios' para 'Clinica hospitalizados' (despachos en farmacia)"
            Top             =   6840
            Width           =   255
         End
         Begin VB.CommandButton Command4 
            Caption         =   "New"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   12420
            TabIndex        =   9
            ToolTipText     =   "New"
            Top             =   300
            Width           =   435
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Del"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   12420
            TabIndex        =   8
            Top             =   570
            Width           =   435
         End
         Begin VB.TextBox txtGalenHos 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   90
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   7020
            Width           =   12255
         End
         Begin VB.CommandButton Command2 
            Caption         =   "..."
            Height          =   225
            Left            =   5850
            TabIndex        =   6
            ToolTipText     =   "Actualiza Precio Venta Farmacia para los nuevos Tarifarios"
            Top             =   6780
            Width           =   345
         End
         Begin VB.CommandButton Command5 
            Caption         =   "..."
            Height          =   195
            Left            =   240
            TabIndex        =   5
            ToolTipText     =   "DesEncriptar texto"
            Top             =   6870
            Width           =   255
         End
         Begin VB.CommandButton Command3 
            Caption         =   "..."
            Height          =   195
            Left            =   1770
            TabIndex        =   4
            ToolTipText     =   "Total Consumo Servicios, segun Nro Cuenta"
            Top             =   6840
            Width           =   255
         End
         Begin VB.CommandButton Command6 
            Caption         =   "..."
            Height          =   195
            Left            =   2160
            TabIndex        =   3
            ToolTipText     =   "Total Consumo Farmacia, segun Nro Cuenta"
            Top             =   6840
            Width           =   255
         End
         Begin VB.CommandButton Command7 
            Caption         =   "..."
            Height          =   195
            Left            =   2790
            TabIndex        =   2
            ToolTipText     =   "Cambia Nro Historia Clinica (ARCHIVO CLINICO)"
            Top             =   6840
            Width           =   255
         End
         Begin MSDataGridLib.DataGrid grdGalenHos 
            Height          =   6585
            Left            =   120
            TabIndex        =   10
            Top             =   300
            Width           =   12225
            _ExtentX        =   21564
            _ExtentY        =   11615
            _Version        =   393216
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin MSDataGridLib.DataGrid grdFacturacionBienesFinanciamiento 
         Height          =   1725
         Left            =   -74940
         TabIndex        =   13
         Top             =   2400
         Width           =   12225
         _ExtentX        =   21564
         _ExtentY        =   3043
         _Version        =   393216
         BackColor       =   8454143
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Detalle Seguros (FacturacionBienesFinanciamientos)"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid grdFacturacionServicioFinanciamientos 
         Height          =   1725
         Left            =   -74970
         TabIndex        =   25
         Top             =   2220
         Width           =   12435
         _ExtentX        =   21934
         _ExtentY        =   3043
         _Version        =   393216
         BackColor       =   8454143
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Detalle Seguros (FacturacionServicioFinanciamientos)"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   345
         Left            =   -74790
         TabIndex        =   52
         Top             =   8490
         Width           =   12525
         _ExtentX        =   22093
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblProcesando 
         Caption         =   "........"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   66
         Top             =   8250
         Width           =   7635
      End
      Begin VB.Label Label2 
         Caption         =   "Nro Cuenta:"
         Height          =   255
         Left            =   -74910
         TabIndex        =   26
         Top             =   8730
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Nro Cuenta:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   11
         Top             =   8760
         Width           =   975
      End
   End
End
Attribute VB_Name = "ComprobantePagoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wrs_Gal As New ADODB.Recordset
Dim oRsFarmMovimientoVentas As New ADODB.Recordset
Dim oRsCajaComprobantePago As New ADODB.Recordset
Dim oRsFacturacionBienesFinanciamiento As New ADODB.Recordset
Dim oRsFactOrdenesBienes As New ADODB.Recordset
Dim oRsFacturacionBienesPagos As New ADODB.Recordset
Dim oRsFactOrdenServicio As New ADODB.Recordset
Dim oRsCajaComprobantePagoS As New ADODB.Recordset
Dim oRsFacturacionServicioFinanciamientos As New ADODB.Recordset
Dim oRsFactOrdenServicioPagos As New ADODB.Recordset
Dim oRsFacturacionServicioPagos As New ADODB.Recordset
Dim lcSql As String
Dim oRsUltCodigo As Long
Const lnIdUsuario As Long = 738
Const lnIdTipoFinanciamiento As Long = 1
Const lnIdFuenteFinanciamiento As Long = 1
Const ln2020 As Long = 9999999
Dim mo_conexion As ADODB.Connection
Dim lnErrCA As Long
Dim ml_Errores As String


Private Sub cmdHistorias_Click()
    If MsgBox("Está seguro ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       Dim oConexHBT As New Connection
       Dim oConexion As New Connection
       Dim oRsTmpHBT1 As New Recordset
       Dim oRsTmpHBT2 As New Recordset
       Dim oRsTmpHBT3 As New Recordset
       Dim oRsTmp1 As New Recordset
       Dim oRsTmp2 As New Recordset
       Dim oRsTmp3 As New Recordset
       Dim oDOMovimientoHistoriaClinica  As New DOMovimientoHistoriaClinica
       Dim oMovimientosHistoriaClinica As New MovimientosHistoriaClinica
       Dim lcSql As String, lnCant As Long, lnTotal As Long
       Dim lnUltimoId As Long
       Dim ms_MensajeError As String
       On Error GoTo Terminar
       Me.MousePointer = 11
       ms_MensajeError = ""
       oConexHBT.Open "dsn=" & txtOdbc.Text
       oConexion.Open sighcomun.CadenaConexion
       oConexion.BeginTrans
       '
       Set oMovimientosHistoriaClinica.Conexion = oConexion
       Set mo_conexion = oConexion
       '
       lnUltimoId = 0
       lcSql = "select * from MovimientosHistoriaClinica order by idMovimiento desc"
       oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
       If oRsTmp1.RecordCount > 0 Then
           lnUltimoId = oRsTmp1.Fields!IdMovimiento
       End If
       oRsTmp1.Close
       lcSql = "select * from MovimientosHistoriaClinica order  by idMovimiento"
       oRsTmpHBT1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
       lnTotal = oRsTmpHBT1.RecordCount
       If lnTotal > 0 Then
          ProgressBar1.Min = 0
          ProgressBar1.Max = lnTotal
          lnCant = 1
          oRsTmpHBT1.MoveFirst
          Do While Not oRsTmpHBT1.EOF
'If lnCant > 100 Then
'Exit Do
'End If
            ProgressBar1.Value = lnCant
            lnCant = lnCant + 1
            'ProgramacionMedica
            With oDOMovimientoHistoriaClinica
                .FechaMovimiento = oRsTmpHBT1.Fields!FechaMovimiento
                '.idAtencion = oRsTmpHBT1.Fields!
                .IdEmpleadoArchivo = oRsTmpHBT1.Fields!IdEmpleadoArchivo
                .IdEmpleadoRecepcion = oRsTmpHBT1.Fields!IdEmpleadoRecepcion
                .IdEmpleadoTransporte = oRsTmpHBT1.Fields!IdEmpleadoTransporte
                .IdGrupoMovimiento = oRsTmpHBT1.Fields!IdGrupoMovimiento
                .IdMotivo = oRsTmpHBT1.Fields!IdMotivo
                .IdMovimiento = oRsTmpHBT1.Fields!IdMovimiento
                .idPaciente = oRsTmpHBT1.Fields!idPaciente
                .idServicioDestino = IIf(IsNull(oRsTmpHBT1.Fields!idServicioDestino), 0, oRsTmpHBT1.Fields!idServicioDestino)
                .IdServicioOrigen = IIf(IsNull(oRsTmpHBT1.Fields!IdServicioOrigen), 0, oRsTmpHBT1.Fields!IdServicioOrigen)
                .IdUsuarioAuditoria = lnIdUsuario
                .NroFolios = IIf(IsNull(oRsTmpHBT1.Fields!NroFolios), 0, oRsTmpHBT1.Fields!NroFolios)
                .Observacion = IIf(IsNull(oRsTmpHBT1.Fields!Observacion), "", oRsTmpHBT1.Fields!Observacion)
            End With
            If lnUltimoId < oDOMovimientoHistoriaClinica.IdMovimiento Then
                If Not InsertarDebbMovimientoHistoriaClinica(oDOMovimientoHistoriaClinica) Then
                      GoTo Terminar
                End If
            ElseIf Me.chkHistorias.Value = 1 Then
                If Not oMovimientosHistoriaClinica.Modificar(oDOMovimientoHistoriaClinica) Then
                     ms_MensajeError = oMovimientosHistoriaClinica.MensajeError: GoTo Terminar
                End If
            End If
            '
            oRsTmpHBT1.MoveNext
          Loop
       End If
       oRsTmpHBT1.Close
       '
       oConexion.CommitTrans
       Me.MousePointer = 1
       Unload Me
    End If
    Exit Sub
            
Terminar:
    oConexion.RollbackTrans
    MsgBox ms_MensajeError
    Me.MousePointer = 1
    Resume

End Sub

Private Sub cmdPacientes_Click()
     If MsgBox("Está seguro ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       Dim oConexHBT As New Connection
       Dim oConexion As New Connection
       Dim oRsTmpHBT1 As New Recordset
       Dim oRsTmpHBT2 As New Recordset
       Dim oRsTmpHBT3 As New Recordset
       Dim oRsTmp1 As New Recordset
       Dim oRsTmp2 As New Recordset
       Dim oRsTmp3 As New Recordset
       Dim oPaciente As New Pacientes
       Dim oDOPaciente As New doPaciente
       Dim oDOHistoria As New DOHistoriaClinica
       Dim oHistoria As New HistoriasClinicas
       Dim lcSql As String, lnCant As Long, lnTotal As Long
       Dim lnUltimoId As Long
       Dim ms_MensajeError As String
       Dim oExcel As Excel.Application
       Dim oSheet As Excel.Worksheet
       Dim j As Integer
       On Error GoTo Terminar
       Me.MousePointer = 11
       ms_MensajeError = ""
       ml_Errores = ""
       oConexHBT.Open "dsn=" & txtOdbc.Text
       oConexion.Open sighcomun.CadenaConexion
       oConexion.BeginTrans
       '
       Set oPaciente.Conexion = oConexion
       Set mo_conexion = oConexion
       Set oHistoria.Conexion = oConexion
       '
       Set oExcel = New Excel.Application
       oExcel.Visible = True
       oExcel.Workbooks.Add
       Set oSheet = oExcel.ActiveSheet
       oSheet.Cells(1, 1).Value = "Error"
       oSheet.Cells(1, 6).Value = "Observación"
       j = 3
       '
       lnUltimoId = 0
       lcSql = "select * from Pacientes order by idPaciente desc"
       oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
       If oRsTmp1.RecordCount > 0 Then
           lnUltimoId = oRsTmp1.Fields!idPaciente
       End If
       oRsTmp1.Close
       If Me.optPacienteAdd.Value = True Then
          lcSql = "select * from Pacientes where idPaciente>" & lnUltimoId & " order by idPaciente"
       Else
          lcSql = "select * from Pacientes order by idPaciente"
       End If
       oRsTmpHBT1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
       lnTotal = oRsTmpHBT1.RecordCount
       If lnTotal > 0 Then
          ProgressBar1.Min = 0
          ProgressBar1.Max = lnTotal
          lnCant = 1
          oRsTmpHBT1.MoveFirst
          Do While Not oRsTmpHBT1.EOF
             ProgressBar1.Value = lnCant
             lnCant = lnCant + 1

             'Tabla: Pacientes
             oDOPaciente.apellidoMaterno = oRsTmpHBT1.Fields!apellidoMaterno
             oDOPaciente.apellidoPaterno = oRsTmpHBT1.Fields!apellidoPaterno
             oDOPaciente.Autogenerado = oRsTmpHBT1.Fields!Autogenerado
             oDOPaciente.DireccionDomicilio = IIf(IsNull(oRsTmpHBT1.Fields!DireccionDomicilio), "", oRsTmpHBT1.Fields!DireccionDomicilio)
             oDOPaciente.FechaNacimiento = IIf(IsNull(oRsTmpHBT1.Fields!FechaNacimiento), 0, oRsTmpHBT1.Fields!FechaNacimiento)
             oDOPaciente.IdCentroPobladoDomicilio = IIf(IsNull(oRsTmpHBT1.Fields!IdCentroPobladoDomicilio), 0, oRsTmpHBT1.Fields!IdCentroPobladoDomicilio)
             oDOPaciente.IdCentroPobladoNacimiento = IIf(IsNull(oRsTmpHBT1.Fields!IdCentroPobladoNacimiento), 0, oRsTmpHBT1.Fields!IdCentroPobladoNacimiento)
             oDOPaciente.IdCentroPobladoProcedencia = IIf(IsNull(oRsTmpHBT1.Fields!IdCentroPobladoProcedencia), 0, oRsTmpHBT1.Fields!IdCentroPobladoProcedencia)
             oDOPaciente.IdDistritoDomicilio = IIf(IsNull(oRsTmpHBT1.Fields!IdDistritoDomicilio), 0, oRsTmpHBT1.Fields!IdDistritoDomicilio)
             oDOPaciente.IdDistritoNacimiento = IIf(IsNull(oRsTmpHBT1.Fields!IdDistritoNacimiento), 0, oRsTmpHBT1.Fields!IdDistritoNacimiento)
             oDOPaciente.IdDistritoProcedencia = IIf(IsNull(oRsTmpHBT1.Fields!IdDistritoProcedencia), 0, oRsTmpHBT1.Fields!IdDistritoProcedencia)
             oDOPaciente.IdDocIdentidad = IIf(IsNull(oRsTmpHBT1.Fields!IdDocIdentidad), 0, oRsTmpHBT1.Fields!IdDocIdentidad)
             oDOPaciente.IdEstadoCivil = IIf(IsNull(oRsTmpHBT1.Fields!IdEstadoCivil), 0, oRsTmpHBT1.Fields!IdEstadoCivil)
             oDOPaciente.IdGradoInstruccion = IIf(IsNull(oRsTmpHBT1.Fields!IdGradoInstruccion), 0, oRsTmpHBT1.Fields!IdGradoInstruccion)
             oDOPaciente.idPaciente = oRsTmpHBT1.Fields!idPaciente
             oDOPaciente.IdPaisDomicilio = IIf(IsNull(oRsTmpHBT1.Fields!IdPaisDomicilio), 0, oRsTmpHBT1.Fields!IdPaisDomicilio)
             oDOPaciente.IdPaisNacimiento = IIf(IsNull(oRsTmpHBT1.Fields!IdPaisNacimiento), 0, oRsTmpHBT1.Fields!IdPaisNacimiento)
             oDOPaciente.IdPaisProcedencia = IIf(IsNull(oRsTmpHBT1.Fields!IdPaisProcedencia), 0, oRsTmpHBT1.Fields!IdPaisProcedencia)
             oDOPaciente.IdProcedencia = IIf(IsNull(oRsTmpHBT1.Fields!IdProcedencia), 0, oRsTmpHBT1.Fields!IdProcedencia)
             oDOPaciente.IdTipoNumeracion = IIf(IsNull(oRsTmpHBT1.Fields!IdTipoNumeracion), 0, oRsTmpHBT1.Fields!IdTipoNumeracion)
             oDOPaciente.idTipoOcupacion = IIf(IsNull(oRsTmpHBT1.Fields!idTipoOcupacion), 0, oRsTmpHBT1.Fields!idTipoOcupacion)
             oDOPaciente.idTipoSexo = IIf(IsNull(oRsTmpHBT1.Fields!idTipoSexo), 1, oRsTmpHBT1.Fields!idTipoSexo)
             oDOPaciente.IdUsuarioAuditoria = lnIdUsuario
             oDOPaciente.NombreMadre = IIf(IsNull(oRsTmpHBT1.Fields!NombreMadre), "", oRsTmpHBT1.Fields!NombreMadre)
             oDOPaciente.NombrePadre = IIf(IsNull(oRsTmpHBT1.Fields!NombrePadre), "", oRsTmpHBT1.Fields!NombrePadre)
             oDOPaciente.NroDocumento = IIf(IsNull(oRsTmpHBT1.Fields!NroDocumento), "", oRsTmpHBT1.Fields!NroDocumento)
             oDOPaciente.NroHistoriaClinica = IIf(IsNull(oRsTmpHBT1.Fields!NroHistoriaClinica), 0, oRsTmpHBT1.Fields!NroHistoriaClinica)
             oDOPaciente.Observacion = IIf(IsNull(oRsTmpHBT1.Fields!Observacion), "", oRsTmpHBT1.Fields!Observacion)
             oDOPaciente.PrimerNombre = oRsTmpHBT1.Fields!PrimerNombre
             oDOPaciente.SegundoNombre = IIf(IsNull(oRsTmpHBT1.Fields!SegundoNombre), "", oRsTmpHBT1.Fields!SegundoNombre)
             oDOPaciente.Telefono = IIf(IsNull(oRsTmpHBT1.Fields!Telefono), "", oRsTmpHBT1.Fields!Telefono)
             oDOPaciente.TercerNombre = IIf(IsNull(oRsTmpHBT1.Fields!TercerNombre), "", oRsTmpHBT1.Fields!TercerNombre)
             If lnUltimoId < oDOPaciente.idPaciente Then
                If Not InsertarTmpPacientesAgregar(oDOPaciente) Then
                      GoTo Terminar
                End If
             ElseIf Me.optPacienteAct.Value = True Then
                If Not oPaciente.Modificar(oDOPaciente, False) Then
                     ms_MensajeError = oPaciente.MensajeError: GoTo Terminar
                End If
             End If
             'Tabla: HistoriasClinicas
             lcSql = "select * from HistoriasClinicas where idPaciente=" & oDOPaciente.idPaciente
             oRsTmpHBT2.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
             If oRsTmpHBT2.RecordCount = 0 Then
'                If oRsTmpHBT1.Fields!IdTipoNumeracion < 4 Then
'                    oSheet.Cells(j, 1).Value = "..<<falta de dato>>...IdPaciente: " & oDOPaciente.idPaciente & " en tabla HistoriasClinicas (Tipo Numeración:" & Trim(Str(oRsTmpHBT1.Fields!IdTipoNumeracion)) & ")"
'                    oSheet.Cells(j, 6).Value = ""
'                    j = j + 1
'
'                End If
                If oRsTmpHBT1.Fields!IdTipoNumeracion < 4 Then
                    oDOHistoria.FechaCreacion = CDate("01/01/2010")
                    oDOHistoria.FechaPasoAPasivo = 0
                    oDOHistoria.IdEstadoHistoria = 1
                    oDOHistoria.idPaciente = oDOPaciente.idPaciente
                    oDOHistoria.IdTipoHistoria = 1
                    oDOHistoria.IdTipoNumeracion = oDOPaciente.IdTipoNumeracion
                    oDOHistoria.IdTipoNumeracionAnterior = oDOPaciente.IdTipoNumeracion
                    oDOHistoria.IdUsuarioAuditoria = lnIdUsuario
                    oDOHistoria.NroHistoriaClinica = oDOPaciente.NroHistoriaClinica
                    oDOHistoria.NroHistoriaClinicaAnterior = 0
                    If Not InsertarDebbHistorias(oDOHistoria) Then
                       If Val(Left(ml_Errores, 11)) = -2147217873 Then
                            oDOHistoria.NroHistoriaClinica = 2000000 + oDOPaciente.NroHistoriaClinica
                            oDOHistoria.NroHistoriaClinicaAnterior = oDOPaciente.NroHistoriaClinica
                            If Not InsertarDebbHistorias(oDOHistoria) Then
                               GoTo Terminar
                            End If
                            oDOPaciente.NroHistoriaClinica = 2000000 + oDOPaciente.NroHistoriaClinica
                            If Not oPaciente.Modificar(oDOPaciente, False) Then
                                 ms_MensajeError = oPaciente.MensajeError: GoTo Terminar
                            End If
                            oSheet.Cells(j, 1).Value = oDOPaciente.NroHistoriaClinica & "..<<historia duplicada>>...IdPaciente: " & oDOPaciente.idPaciente & "....tipo numero=" & oDOPaciente.IdTipoNumeracion
                            oSheet.Cells(j, 6).Value = ""
                            j = j + 1
                       Else
                         GoTo Terminar
                       End If
                    End If
                End If
             Else
                oDOHistoria.FechaCreacion = oRsTmpHBT2.Fields!FechaCreacion
                oDOHistoria.FechaPasoAPasivo = IIf(IsNull(oRsTmpHBT2.Fields!FechaPasoAPasivo), 0, oRsTmpHBT2.Fields!FechaPasoAPasivo)
                oDOHistoria.IdEstadoHistoria = oRsTmpHBT2.Fields!IdEstadoHistoria
                oDOHistoria.idPaciente = oRsTmpHBT2.Fields!idPaciente
                oDOHistoria.IdTipoHistoria = IIf(IsNull(oRsTmpHBT2.Fields!IdTipoHistoria), 0, oRsTmpHBT2.Fields!IdTipoHistoria)
                oDOHistoria.IdTipoNumeracion = oRsTmpHBT2.Fields!IdTipoNumeracion
                oDOHistoria.IdTipoNumeracionAnterior = IIf(IsNull(oRsTmpHBT2.Fields!IdTipoNumeracionAnterior), 0, oRsTmpHBT2.Fields!IdTipoNumeracionAnterior)
                oDOHistoria.IdUsuarioAuditoria = lnIdUsuario
                oDOHistoria.NroHistoriaClinica = oRsTmpHBT2.Fields!NroHistoriaClinica
                oDOHistoria.NroHistoriaClinicaAnterior = IIf(IsNull(oRsTmpHBT2.Fields!NroHistoriaClinicaAnterior), 0, oRsTmpHBT2.Fields!NroHistoriaClinicaAnterior)
                If lnUltimoId < oDOPaciente.idPaciente Then
                    If Not InsertarDebbHistorias(oDOHistoria) Then
                          GoTo Terminar
                    End If
                Else
                    If Not oHistoria.Modificar(oDOHistoria) Then
                        ms_MensajeError = oHistoria.MensajeError: GoTo Terminar
                    End If
                End If
             End If
             oRsTmpHBT2.Close
             '
             oRsTmpHBT1.MoveNext
          Loop
       End If
       oRsTmpHBT1.Close
       '
       oConexion.CommitTrans
       Me.MousePointer = 1
'       oSheet.SaveAs "c:\estructura.xls"
 '      MsgBox "Se grabó c:\estructura.xls"
       Unload Me
    End If
    Exit Sub
            
Terminar:
    oConexion.RollbackTrans
    MsgBox ms_MensajeError & ml_Errores
    Me.MousePointer = 1
    oSheet.SaveAs "c:\estructura.xls"
    MsgBox "Se grabó c:\estructura.xls"
    Resume
End Sub

Private Sub cmdProcesaAtenciones_Click()
    If MsgBox("Está seguro ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       Dim oConexHBT As New Connection
       Dim oConexion As New Connection
       Dim oRsTmpHBT1 As New Recordset
       Dim oRsTmpHBT2 As New Recordset
       Dim oRsTmpHBT3 As New Recordset
       Dim oRsTmp1 As New Recordset
       Dim oRsTmp2 As New Recordset
       Dim oRsTmp3 As New Recordset
       Dim oDOAtencion As New DOAtencion
       Dim oAtenciones As New Atenciones
       Dim oDOCuentaAtencion As New DOCuentaAtencion
       Dim oCuentaAtencion As New CuentasAtencion
       Dim oDOAtencionDiagnostico As New DOAtencionDiagnostico
       Dim oAtencionesDiagnosticos As New AtencionesDiagnosticos
       Dim oAtencionesEmergencia As New AtencionesEmergencia
       Dim oDOAtencionEmergencia As New DOAtencionEmergencia
       Dim oAtencionesEstanciaHosp As New AtencionesEstanciaHosp
       Dim oDOEstanciaHospitalaria As New DOEstanciaHospitalaria
       Dim oAtencionesNacimientos As New AtencionesNacimientos
       Dim oDOAtencionNacimiento As New DOAtencionNacimiento
       Dim oCitas As New Citas
       Dim oDoCita As New DOCita
       Dim lcSql As String, lnCant As Long, lnTotal As Long, lcEstoyEn As String
       Dim ms_MensajeError As String
       Dim oExcel As Excel.Application
       Dim oSheet As Excel.Worksheet
       Dim j As Integer
       Dim lnIdCuentaAtencion2020 As Long
       On Error GoTo Terminar
       Me.MousePointer = 11
       ms_MensajeError = ""
       oConexHBT.Open "dsn=" & txtOdbc.Text
       oConexion.Open sighcomun.CadenaConexion
       oConexion.BeginTrans
       '
       Set oAtenciones.Conexion = oConexion
       Set oCuentaAtencion.Conexion = oConexion
       Set oAtencionesDiagnosticos.Conexion = oConexion
       Set oAtencionesEmergencia.Conexion = oConexion
       Set oAtencionesEstanciaHosp.Conexion = oConexion
       Set oAtencionesNacimientos.Conexion = oConexion
       Set oCitas.Conexion = oConexion
       Set mo_conexion = oConexion
       '
       lcEstoyEn = ""
       lcSql = "select * from atenciones where FechaIngreso>='" & txtFinicial.Text & "' and fechaIngreso<='" & txtFfinal.Text & "' order by FechaIngreso"
       oRsTmpHBT1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
       lnTotal = oRsTmpHBT1.RecordCount
       If lnTotal > 0 Then
          ProgressBar1.Min = 0
          lnCant = 1
          'Elimina datos anteriores
          lcEstoyEn = "Elimina datos anteriores"
          Me.lblProcesando.Caption = "...Eliminando Datos"
          lcSql = "select * from atenciones where FechaIngreso>='" & txtFinicial.Text & "' and fechaIngreso<='" & txtFfinal.Text & "'"
          oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
         
          If oRsTmp1.RecordCount > 0 Then
              ProgressBar1.Max = oRsTmp1.RecordCount
             oRsTmp1.MoveFirst
             Do While Not oRsTmp1.EOF

                ProgressBar1.Value = lnCant
                lnCant = lnCant + 1
                lcSql = "delete from FacturacionCuentasAtencion where idCuentaAtencion=" & oRsTmp1.Fields!idCuentaAtencion
                oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                lcSql = "delete from AtencionesDiagnosticos where idAtencion=" & oRsTmp1.Fields!idAtencion
                oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                lcSql = "delete from AtencionesEmergencia where idAtencion=" & oRsTmp1.Fields!idAtencion
                oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                lcSql = "delete from AtencionesEstanciaHospitalaria where idAtencion=" & oRsTmp1.Fields!idAtencion
                oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                lcSql = "delete from AtencionesNacimientos where idAtencion=" & oRsTmp1.Fields!idAtencion
                oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                lcSql = "delete from Citas where idAtencion=" & oRsTmp1.Fields!idAtencion
                oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                oRsTmp1.Delete
                oRsTmp1.Update
                oRsTmp1.MoveNext
             Loop
          End If
          '
          Me.lblProcesando.Caption = "...Procesando Datos"
          ProgressBar1.Max = lnTotal
          lnCant = 1
          '
          Set oExcel = New Excel.Application
          oExcel.Visible = True
          oExcel.Workbooks.Add
          Set oSheet = oExcel.ActiveSheet
          oSheet.Cells(1, 1).Value = "Error"
          oSheet.Cells(1, 6).Value = "Observación"
          j = 3
          '
          oRsTmpHBT1.MoveFirst
          Do While Not oRsTmpHBT1.EOF

             ProgressBar1.Value = lnCant
             lnCant = lnCant + 1
             lcSql = "select idPaciente from Pacientes where idPaciente=" & oRsTmpHBT1.Fields!idPaciente
             oRsTmp3.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
             If oRsTmp3.RecordCount = 0 Then
                oSheet.Cells(j, 1).Value = "..<<falta de dato>>...IdPaciente: " & oRsTmpHBT1.Fields!idPaciente & " en tabla Pacientes (idAtencion=" & Trim(Str(oRsTmpHBT1.Fields!idAtencion)) & ") (idTipoServicio: " & Trim(Str(oRsTmpHBT1.Fields!IdTipoServicio)) & ")"
                oSheet.Cells(j, 6).Value = ""
                j = j + 1
                oRsTmp3.Close
             Else
                 oRsTmp3.Close
                 'tabla: Cuentas de Atencion
                 lcEstoyEn = "Cuentas de Atencion"
                 lcSql = "select * from FacturacionCuentasAtencion where idCuentaAtencion=" & oRsTmpHBT1.Fields!idCuentaAtencion
                 oRsTmpHBT2.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
                 If oRsTmpHBT2.RecordCount = 0 Then
'                    oSheet.Cells(j, 1).Value = "..<<falta de dato>>...IdCuentaAtencion: " & oRsTmpHBT1.Fields!idCuentaAtencion & " en tabla FacturacionCuentasAtencion (idTipoServicio: " & Trim(Str(oRsTmpHBT1.Fields!IdTipoServicio)) & ")"
'                    oSheet.Cells(j, 6).Value = ""
'                    j = j + 1
                    oDOCuentaAtencion.FechaApertura = oRsTmpHBT1.Fields!FechaIngreso
                    oDOCuentaAtencion.FechaCierre = 0
                    oDOCuentaAtencion.FechaCreacion = oRsTmpHBT1.Fields!FechaIngreso
                    oDOCuentaAtencion.HoraApertura = oRsTmpHBT1.Fields!HoraIngreso
                    oDOCuentaAtencion.HoraCierre = ""
                    oDOCuentaAtencion.idCuentaAtencion = oRsTmpHBT1.Fields!idAtencion
                    oDOCuentaAtencion.idEstado = 0
                    oDOCuentaAtencion.idPaciente = oRsTmpHBT1.Fields!idPaciente
                    oDOCuentaAtencion.IdUsuarioAuditoria = lnIdUsuario
                    oDOCuentaAtencion.TotalAsegurado = 0
                    oDOCuentaAtencion.TotalExonerado = 0
                    oDOCuentaAtencion.TotalPagado = 0
                    oDOCuentaAtencion.TotalPorPagar = 0
                 Else
                    oDOCuentaAtencion.FechaApertura = IIf(IsNull(oRsTmpHBT2.Fields!FechaApertura), 0, oRsTmpHBT2.Fields!FechaApertura)
                    oDOCuentaAtencion.FechaCierre = IIf(IsNull(oRsTmpHBT2.Fields!FechaCierre), 0, oRsTmpHBT2.Fields!FechaCierre)
                    oDOCuentaAtencion.FechaCreacion = IIf(IsNull(oRsTmpHBT2.Fields!FechaCreacion), 0, oRsTmpHBT2.Fields!FechaCreacion)
                    oDOCuentaAtencion.HoraApertura = IIf(IsNull(oRsTmpHBT2.Fields!HoraApertura), "", oRsTmpHBT2.Fields!HoraApertura)
                    oDOCuentaAtencion.HoraCierre = IIf(IsNull(oRsTmpHBT2.Fields!HoraCierre), "", oRsTmpHBT2.Fields!HoraCierre)
                    oDOCuentaAtencion.idCuentaAtencion = oRsTmpHBT1.Fields!idAtencion
                    oDOCuentaAtencion.idEstado = IIf(IsNull(oRsTmpHBT2.Fields!idEstado), 0, oRsTmpHBT2.Fields!idEstado)
                    oDOCuentaAtencion.idPaciente = oRsTmpHBT2.Fields!idPaciente
                    oDOCuentaAtencion.IdUsuarioAuditoria = lnIdUsuario
                    oDOCuentaAtencion.TotalAsegurado = IIf(IsNull(oRsTmpHBT2.Fields!TotalAsegurado), 0, oRsTmpHBT2.Fields!TotalAsegurado)
                    oDOCuentaAtencion.TotalExonerado = IIf(IsNull(oRsTmpHBT2.Fields!TotalExonerado), 0, oRsTmpHBT2.Fields!TotalExonerado)
                    oDOCuentaAtencion.TotalPagado = IIf(IsNull(oRsTmpHBT2.Fields!TotalPagado), 0, oRsTmpHBT2.Fields!TotalPagado)
                    oDOCuentaAtencion.TotalPorPagar = IIf(IsNull(oRsTmpHBT2.Fields!TotalPorPagar), 0, oRsTmpHBT2.Fields!TotalPorPagar)
                End If
                lnErrCA = 0
                If Not InsertarDebbCuentaAtencion(oDOCuentaAtencion) Then
                        oSheet.Cells(j, 1).Value = "..<<ya existe>>...IdCuentaAtencion: " & oRsTmpHBT1.Fields!idCuentaAtencion & " en tabla FacturacionCuentasAtencion (idAtencion=" & Trim(Str(oRsTmpHBT1.Fields!idAtencion)) & ") (FechaIngreso: " & oRsTmpHBT1.Fields!FechaIngreso & ")"
                        oSheet.Cells(j, 6).Value = ""
                        j = j + 1
                        GoTo Terminar
                End If
                If lnErrCA = 0 Then
                    'tabla: Atenciones
                    lcEstoyEn = "Atencion"
                    With oDOAtencion
                        .DireccionDomicilio = IIf(IsNull(oRsTmpHBT1.Fields!DireccionDomicilio), "", oRsTmpHBT1.Fields!DireccionDomicilio)
                        .Edad = IIf(IsNull(oRsTmpHBT1.Fields!Edad), 0, oRsTmpHBT1.Fields!Edad)
                        .EsPacienteExterno = False
                        .FechaEgreso = IIf(IsNull(oRsTmpHBT1.Fields!FechaEgreso), 0, oRsTmpHBT1.Fields!FechaEgreso)
                        .FechaEgresoAdministrativo = IIf(IsNull(oRsTmpHBT1.Fields!FechaEgresoAdministrativo), 0, oRsTmpHBT1.Fields!FechaEgresoAdministrativo)
                        .FechaIngreso = oRsTmpHBT1.Fields!FechaIngreso
                        .HoraEgreso = IIf(IsNull(oRsTmpHBT1.Fields!HoraEgreso), 0, oRsTmpHBT1.Fields!HoraEgreso)
                        .HoraEgresoAdministrativo = IIf(IsNull(oRsTmpHBT1.Fields!HoraEgresoAdministrativo), 0, oRsTmpHBT1.Fields!HoraEgresoAdministrativo)
                        .HoraIngreso = oRsTmpHBT1.Fields!HoraIngreso
                        .HuboInfeccionIntraHospitalaria = IIf(IsNull(oRsTmpHBT1.Fields!HuboInfeccionIntraHospitalaria), 0, oRsTmpHBT1.Fields!HuboInfeccionIntraHospitalaria)
                        .idAtencion = oRsTmpHBT1.Fields!idAtencion
                        If IsNull(oRsTmpHBT1.Fields!IdCamaEgreso) Then
                          If oDOAtencion.IdTipoServicio = 3 Then
                             .IdCamaEgreso = .IdCamaIngreso
                          End If
                        Else
                          .IdCamaEgreso = oRsTmpHBT1.Fields!IdCamaEgreso
                        End If
                        .IdCamaIngreso = IIf(IsNull(oRsTmpHBT1.Fields!IdCamaIngreso), 0, oRsTmpHBT1.Fields!IdCamaIngreso)
                        .IdCondicionAlta = IIf(IsNull(oRsTmpHBT1.Fields!IdCondicionAlta), 0, oRsTmpHBT1.Fields!IdCondicionAlta)
                        .idCuentaAtencion = oRsTmpHBT1.Fields!idAtencion
                        .IdDestinoAtencion = IIf(IsNull(oRsTmpHBT1.Fields!IdDestinoAtencion), 0, oRsTmpHBT1.Fields!IdDestinoAtencion)
                        .IdEspecialidadMedico = IIf(IsNull(oRsTmpHBT1.Fields!IdEspecialidadMedico), 0, oRsTmpHBT1.Fields!IdEspecialidadMedico)
                        .IdEstablecimientoDestino = IIf(IsNull(oRsTmpHBT1.Fields!IdEstablecimientoDestino), 0, oRsTmpHBT1.Fields!IdEstablecimientoDestino)
                        .IdEstablecimientoNoMinsaDestino = IIf(IsNull(oRsTmpHBT1.Fields!IdEstablecimientoNoMinsaDestino), 0, oRsTmpHBT1.Fields!IdEstablecimientoNoMinsaDestino)
                        .IdEstablecimientoNoMinsaOrigen = IIf(IsNull(oRsTmpHBT1.Fields!IdEstablecimientoNoMinsaOrigen), 0, oRsTmpHBT1.Fields!IdEstablecimientoNoMinsaOrigen)
                        .IdEstablecimientoOrigen = IIf(IsNull(oRsTmpHBT1.Fields!IdEstablecimientoOrigen), 0, oRsTmpHBT1.Fields!IdEstablecimientoOrigen)
                        .IdEstadoAtencion = 1  'registrado
                        .IdFormaPago = lnIdTipoFinanciamiento
                        .IdFuenteFinanciamiento = lnIdFuenteFinanciamiento
                        If IsNull(oRsTmpHBT1.Fields!IdMedicoEgreso) Then
                           .IdMedicoEgreso = IIf(IsNull(oRsTmpHBT1.Fields!IdServicioEgreso), 0, oRsTmpHBT1.Fields!IdMedicoIngreso)
                        Else
                           .IdMedicoEgreso = oRsTmpHBT1.Fields!IdMedicoEgreso
                        End If
                        .IdMedicoIngreso = oRsTmpHBT1.Fields!IdMedicoIngreso
                        .IdMedicoRespNacimiento = IIf(IsNull(oRsTmpHBT1.Fields!IdMedicoRespNacimiento), 0, oRsTmpHBT1.Fields!IdMedicoRespNacimiento)
                        .IdOrigenAtencion = IIf(IsNull(oRsTmpHBT1.Fields!IdOrigenAtencion), 0, oRsTmpHBT1.Fields!IdOrigenAtencion)
                        .idPaciente = oRsTmpHBT1.Fields!idPaciente
                        .IdServicioEgreso = IIf(IsNull(oRsTmpHBT1.Fields!IdServicioEgreso), 0, oRsTmpHBT1.Fields!IdServicioEgreso)
                        .IdServicioIngreso = IIf(IsNull(oRsTmpHBT1.Fields!IdServicioIngreso), 0, oRsTmpHBT1.Fields!IdServicioIngreso)
                        .IdTipoAlta = IIf(IsNull(oRsTmpHBT1.Fields!IdTipoAlta), 0, oRsTmpHBT1.Fields!IdTipoAlta)
                        .IdTipoCondicionALEstab = IIf(IsNull(oRsTmpHBT1.Fields!IdTipoCondicionALEstab), 0, oRsTmpHBT1.Fields!IdTipoCondicionALEstab)
                        .IdTipoCondicionAlServicio = IIf(IsNull(oRsTmpHBT1.Fields!IdTipoCondicionAlServicio), 0, oRsTmpHBT1.Fields!IdTipoCondicionAlServicio)
                        .IdTipoEdad = IIf(IsNull(oRsTmpHBT1.Fields!IdTipoEdad), 0, oRsTmpHBT1.Fields!IdTipoEdad)
                        .IdTipoGravedad = IIf(IsNull(oRsTmpHBT1.Fields!IdTipoGravedad), 0, oRsTmpHBT1.Fields!IdTipoGravedad)
                        .IdTipoReferenciaDestino = IIf(IsNull(oRsTmpHBT1.Fields!IdTipoReferenciaDestino), 0, oRsTmpHBT1.Fields!IdTipoReferenciaDestino)
                        .IdTipoReferenciaOrigen = IIf(IsNull(oRsTmpHBT1.Fields!IdTipoReferenciaOrigen), 0, oRsTmpHBT1.Fields!IdTipoReferenciaOrigen)
                        .IdTipoServicio = oRsTmpHBT1.Fields!IdTipoServicio
                        .IdUsuarioAuditoria = lnIdUsuario
                        .NombreAcompaniante = IIf(IsNull(oRsTmpHBT1.Fields!NombreAcompaniante), "", oRsTmpHBT1.Fields!NombreAcompaniante)
                        .NroReferenciaDestino = ""
                        .NroReferenciaOrigen = ""
                        .Observacion = IIf(IsNull(oRsTmpHBT1.Fields!Observacion), "", oRsTmpHBT1.Fields!Observacion)
                        .pisoDomicilio = IIf(IsNull(oRsTmpHBT1.Fields!pisoDomicilio), "", oRsTmpHBT1.Fields!pisoDomicilio)
                        .RecienNacido = IIf(IsNull(oRsTmpHBT1.Fields!RecienNacido), 0, oRsTmpHBT1.Fields!RecienNacido)
                        .TieneNecropsia = IIf(IsNull(oRsTmpHBT1.Fields!TieneNecropsia), 0, oRsTmpHBT1.Fields!TieneNecropsia)
                    End With
                    If Not InsertarDebbAtenciones(oDOAtencion) Then
                          GoTo Terminar
                    End If
                    'AtencionesDiagnosticos
                    lcEstoyEn = "AtencionesDiagnosticos"
                    lcSql = "select * from AtencionesDiagnosticos where idAtencion=" & oRsTmpHBT1.Fields!idAtencion
                    oRsTmpHBT3.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
                    If oRsTmpHBT3.RecordCount > 0 Then
                       oRsTmpHBT3.MoveFirst
                       Do While Not oRsTmpHBT3.EOF
                            With oDOAtencionDiagnostico
                                .idAtencion = oRsTmpHBT3.Fields!idAtencion
                                .IdAtencionDiagnostico = oRsTmpHBT3.Fields!IdAtencionDiagnostico
                                .IdClasificacionDx = IIf(IsNull(oRsTmpHBT3.Fields!IdClasificacionDx), 0, oRsTmpHBT3.Fields!IdClasificacionDx)
                                .IdDiagnostico = IIf(IsNull(oRsTmpHBT3.Fields!IdDiagnostico), 0, oRsTmpHBT3.Fields!IdDiagnostico)
                                .IdSubclasificacionDx = IIf(IsNull(oRsTmpHBT3.Fields!IdSubclasificacionDx), 0, oRsTmpHBT3.Fields!IdSubclasificacionDx)
                                .IdUsuarioAuditoria = lnIdUsuario
                                .labConfHIS = ""
                            End With
                            If Not InsertarDebbAtencionDiagnostico(oDOAtencionDiagnostico) Then
                                 GoTo Terminar
                            End If
                            oRsTmpHBT3.MoveNext
                       Loop
                    End If
                    oRsTmpHBT3.Close
                    If oDOAtencion.IdTipoServicio > 1 Then
                        'atencionesEmergencia
                        lcEstoyEn = "atencionesEmergencia"
                        lcSql = "select * from AtencionesEmergencia where idAtencion=" & oRsTmpHBT1.Fields!idAtencion
                        oRsTmpHBT3.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
                        If oRsTmpHBT3.RecordCount > 0 Then
                           oRsTmpHBT3.MoveFirst
                           Do While Not oRsTmpHBT3.EOF
                                With oDOAtencionEmergencia
                                    .idAtencion = oRsTmpHBT3.Fields!idAtencion
                                    .IdAtencionEmergencia = oRsTmpHBT3.Fields!IdAtencionEmergencia
                                    .IdCausaExternaMorbilidad = IIf(IsNull(oRsTmpHBT3.Fields!IdCausaExternaMorbilidad), 0, oRsTmpHBT3.Fields!IdCausaExternaMorbilidad)
                                    .IdClaseAccidente = IIf(IsNull(oRsTmpHBT3.Fields!IdClaseAccidente), 0, oRsTmpHBT3.Fields!IdClaseAccidente)
                                    .IdGrupoOcupacionalALAB = IIf(IsNull(oRsTmpHBT3.Fields!IdGrupoOcupacionalALAB), 0, oRsTmpHBT3.Fields!IdGrupoOcupacionalALAB)
                                    .IdLugarEvento = IIf(IsNull(oRsTmpHBT3.Fields!IdLugarEvento), 0, oRsTmpHBT3.Fields!IdLugarEvento)
                                    .IdPosicionLesionadoALAB = IIf(IsNull(oRsTmpHBT3.Fields!IdPosicionLesionadoALAB), 0, oRsTmpHBT3.Fields!IdPosicionLesionadoALAB)
                                    .IdRelacionAgresorVictima = IIf(IsNull(oRsTmpHBT3.Fields!IdRelacionAgresorVictima), 0, oRsTmpHBT3.Fields!IdRelacionAgresorVictima)
                                    .IdSeguridad = IIf(IsNull(oRsTmpHBT3.Fields!IdSeguridad), 0, oRsTmpHBT3.Fields!IdSeguridad)
                                    .IdTipoAgenteAGAN = IIf(IsNull(oRsTmpHBT3.Fields!IdTipoAgenteAGAN), 0, oRsTmpHBT3.Fields!IdTipoAgenteAGAN)
                                    .IdTipoEvento = IIf(IsNull(oRsTmpHBT3.Fields!IdTipoEvento), 0, oRsTmpHBT3.Fields!IdTipoEvento)
                                    .IdTipoTransporte = IIf(IsNull(oRsTmpHBT3.Fields!IdTipoTransporte), 0, oRsTmpHBT3.Fields!IdTipoTransporte)
                                    .IdTipoVehiculo = IIf(IsNull(oRsTmpHBT3.Fields!IdTipoVehiculo), 0, oRsTmpHBT3.Fields!IdTipoVehiculo)
                                    .IdUbicacionLesionado = IIf(IsNull(oRsTmpHBT3.Fields!IdUbicacionLesionado), 0, oRsTmpHBT3.Fields!IdUbicacionLesionado)
                                    .IdUsuarioAuditoria = lnIdUsuario
                                End With
                                If Not InsertarDebbAtencionEmergencia(oDOAtencionEmergencia) Then
                                      GoTo Terminar
                                End If
                                oRsTmpHBT3.MoveNext
                           Loop
                        End If
                        oRsTmpHBT3.Close
                        'atencionesEstanciaHospitalaria
                        lcEstoyEn = "atencionesEstanciaHospitalaria"
                        lcSql = "select * from AtencionesEstanciaHospitalaria where idAtencion=" & oRsTmpHBT1.Fields!idAtencion
                        oRsTmpHBT3.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
                        If oRsTmpHBT3.RecordCount > 0 Then
                           oRsTmpHBT3.MoveFirst
                           Do While Not oRsTmpHBT3.EOF
                                With oDOEstanciaHospitalaria
                                    .DiasEstancia = IIf(IsNull(oRsTmpHBT3.Fields!DiasEstancia), 0, oRsTmpHBT3.Fields!DiasEstancia)
                                    .FechaDesocupacion = IIf(IsNull(oRsTmpHBT3.Fields!FechaDesocupacion), 0, oRsTmpHBT3.Fields!FechaDesocupacion)
                                    .FechaOcupacion = IIf(IsNull(oRsTmpHBT3.Fields!FechaOcupacion), 0, oRsTmpHBT3.Fields!FechaOcupacion)
                                    .HoraDesocupacion = IIf(IsNull(oRsTmpHBT3.Fields!HoraDesocupacion), "", oRsTmpHBT3.Fields!HoraDesocupacion)
                                    .HoraOcupacion = IIf(IsNull(oRsTmpHBT3.Fields!HoraOcupacion), "", oRsTmpHBT3.Fields!HoraOcupacion)
                                    .idAtencion = oRsTmpHBT3.Fields!idAtencion
                                    .idCama = IIf(IsNull(oRsTmpHBT3.Fields!idCama), 0, oRsTmpHBT3.Fields!idCama)
                                    .IdEstanciaHospitalaria = oRsTmpHBT3.Fields!IdEstanciaHospitalaria
                                    .IdFacturacionServicio = IIf(IsNull(oRsTmpHBT3.Fields!IdFacturacionServicio), 0, oRsTmpHBT3.Fields!IdFacturacionServicio)
                                    .IdMedicoOrdena = IIf(IsNull(oRsTmpHBT3.Fields!IdMedicoOrdena), 0, oRsTmpHBT3.Fields!IdMedicoOrdena)
                                    .idProducto = 4590
                                    .idServicio = IIf(IsNull(oRsTmpHBT3.Fields!idServicio), 0, oRsTmpHBT3.Fields!idServicio)
                                    .IdUsuarioAuditoria = lnIdUsuario
                                    .LlegoAlServicio = 1
                                    .Secuencia = oRsTmpHBT3.Fields!Secuencia
                                End With
                                If Not InsertarDebbEstanciaHospitalaria(oDOEstanciaHospitalaria) Then
                                     GoTo Terminar
                                End If
                                oRsTmpHBT3.MoveNext
                           Loop
                        End If
                        oRsTmpHBT3.Close
                        'atencionesNacimientos
                        lcEstoyEn = "atencionesNacimientos"
                        lcSql = "select * from AtencionesNacimientos where idAtencion=" & oRsTmpHBT1.Fields!idAtencion
                        oRsTmpHBT3.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
                        If oRsTmpHBT3.RecordCount > 0 Then
                           oRsTmpHBT3.MoveFirst
                           Do While Not oRsTmpHBT3.EOF
                                With oDOAtencionNacimiento
                                    .EdadSemanas = IIf(IsNull(oRsTmpHBT3.Fields!EdadSemanas), 0, oRsTmpHBT3.Fields!EdadSemanas)
                                    .FechaNacimiento = IIf(IsNull(oRsTmpHBT3.Fields!FechaNacimiento), 0, oRsTmpHBT3.Fields!FechaNacimiento)
                                    .idAtencion = oRsTmpHBT3.Fields!idAtencion
                                    .idCondicionRN = IIf(IsNull(oRsTmpHBT3.Fields!idCondicionRN), 0, oRsTmpHBT3.Fields!idCondicionRN)
                                    .idNacimiento = oRsTmpHBT3.Fields!idNacimiento
                                    .idTipoSexo = IIf(IsNull(oRsTmpHBT3.Fields!idTipoSexo), 0, oRsTmpHBT3.Fields!idTipoSexo)
                                    .IdUsuarioAuditoria = lnIdUsuario
                                    .Peso = IIf(IsNull(oRsTmpHBT3.Fields!Peso), 0, oRsTmpHBT3.Fields!Peso)
                                    .Talla = IIf(IsNull(oRsTmpHBT3.Fields!Talla), 0, oRsTmpHBT3.Fields!Talla)
                                End With
                                If Not InsertarDebbAtencionNacimiento(oDOAtencionNacimiento) Then
                                      GoTo Terminar
                                End If
                                oRsTmpHBT3.MoveNext
                           Loop
                        End If
                        oRsTmpHBT3.Close
                    Else
                        'citas
                        lcEstoyEn = "citas"
                        lcSql = "select * from Citas where idAtencion=" & oRsTmpHBT1.Fields!idAtencion
                        oRsTmpHBT3.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
                        If oRsTmpHBT3.RecordCount > 0 Then
                            With oDoCita
                                .Fecha = oRsTmpHBT3.Fields!Fecha
                                .FechaSolicitud = IIf(IsNull(oRsTmpHBT3.Fields!FechaSolicitud), 0, oRsTmpHBT3.Fields!FechaSolicitud)
                                .HoraFin = oRsTmpHBT3.Fields!HoraFin
                                .HoraInicio = oRsTmpHBT3.Fields!HoraInicio
                                .HoraSolicitud = IIf(IsNull(oRsTmpHBT3.Fields!HoraSolicitud), "", oRsTmpHBT3.Fields!HoraSolicitud)
                                .idAtencion = oRsTmpHBT3.Fields!idAtencion
                                .IdCita = oRsTmpHBT3.Fields!IdCita
                                .IdEspecialidad = IIf(IsNull(oRsTmpHBT3.Fields!IdEspecialidad), 0, oRsTmpHBT3.Fields!IdEspecialidad)
                                .IdEstadoCita = IIf(IsNull(oRsTmpHBT3.Fields!IdEstadoCita), 0, oRsTmpHBT3.Fields!IdEstadoCita)
                                .idMedico = oRsTmpHBT3.Fields!idMedico
                                .idPaciente = oRsTmpHBT3.Fields!idPaciente
                                .idProducto = IIf(IsNull(oRsTmpHBT3.Fields!idProducto), 0, oRsTmpHBT3.Fields!idProducto)
                                .IdProgramacion = oRsTmpHBT3.Fields!IdProgramacion
                                .idServicio = oRsTmpHBT3.Fields!idServicio
                                .IdUsuarioAuditoria = lnIdUsuario
                            End With
                            If Not InsertarDebbCita(oDoCita) Then
                                  GoTo Terminar
                            End If
                        End If
                        oRsTmpHBT3.Close
                    End If
                End If
                '
                oRsTmpHBT2.Close
             End If
             '
             oRsTmpHBT1.MoveNext
          Loop
       End If
       oRsTmpHBT1.Close
       '
       lcEstoyEn = "..Todo OK..."
       oConexion.CommitTrans
       Me.MousePointer = 1
       'oSheet.SaveAs "c:\estructura.xls"
       'MsgBox "Se grabó c:\estructura.xls"
       Unload Me
    End If
    Exit Sub
Terminar:
    If ms_MensajeError = "" Then
       ms_MensajeError = Err.Description
    End If
    If MsgBox(ms_MensajeError & Chr(13) & "iva a Grabar en:" & lcEstoyEn & Chr(13) & Chr(13) & "Desea grabar la información hasta el " & Format((oRsTmpHBT1.Fields!FechaIngreso - 1), "dd/mm/yyyy") & " ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       oConexion.CommitTrans
       Me.MousePointer = 1
       Unload Me
    Else
       oConexion.RollbackTrans
    End If
    Me.MousePointer = 1
   ' Resume
End Sub

Private Sub cmdProgramacion_Click()
    If MsgBox("Está seguro ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       Dim oConexHBT As New Connection
       Dim oConexion As New Connection
       Dim oRsTmpHBT1 As New Recordset
       Dim oRsTmpHBT2 As New Recordset
       Dim oRsTmpHBT3 As New Recordset
       Dim oRsTmp1 As New Recordset
       Dim oRsTmp2 As New Recordset
       Dim oRsTmp3 As New Recordset
       Dim oDOProgramacionMedica  As New DOProgramacionMedica
       Dim oProgramacionMedica As New ProgramacionMedica
       Dim lcSql As String, lnCant As Long, lnTotal As Long
       Dim lnUltimoId As Long
       Dim ms_MensajeError As String
       On Error GoTo Terminar
       Me.MousePointer = 11
       ms_MensajeError = ""
       oConexHBT.Open "dsn=" & txtOdbc.Text
       oConexion.Open sighcomun.CadenaConexion
       oConexion.BeginTrans
       '
       Set oProgramacionMedica.Conexion = oConexion
       Set mo_conexion = oConexion
       '
       lnUltimoId = 0
       lcSql = "select * from ProgramacionMedica order by idProgramacion desc"
       oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
       If oRsTmp1.RecordCount > 0 Then
           lnUltimoId = oRsTmp1.Fields!IdProgramacion
       End If
       oRsTmp1.Close
       lcSql = "select * from ProgramacionMedica order  by idProgramacion"
       oRsTmpHBT1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
       lnTotal = oRsTmpHBT1.RecordCount
       If lnTotal > 0 Then
          ProgressBar1.Min = 0
          ProgressBar1.Max = lnTotal
          lnCant = 1
          oRsTmpHBT1.MoveFirst
          Do While Not oRsTmpHBT1.EOF
If lnCant > 100000 Then
Exit Do
End If
            ProgressBar1.Value = lnCant
            lnCant = lnCant + 1
            'ProgramacionMedica
            With oDOProgramacionMedica
                .Color = IIf(IsNull(oRsTmpHBT1.Fields!Color), 0, oRsTmpHBT1.Fields!Color)
                .Descripcion = IIf(IsNull(oRsTmpHBT1.Fields!Descripcion), "", oRsTmpHBT1.Fields!Descripcion)
                .Fecha = oRsTmpHBT1.Fields!Fecha
                .HoraFin = oRsTmpHBT1.Fields!HoraFin
                .HoraInicio = oRsTmpHBT1.Fields!HoraInicio
                .IdDepartamento = IIf(IsNull(oRsTmpHBT1.Fields!IdDepartamento), 0, oRsTmpHBT1.Fields!IdDepartamento)
                .IdEspecialidad = IIf(IsNull(oRsTmpHBT1.Fields!IdEspecialidad), 0, oRsTmpHBT1.Fields!IdEspecialidad)
                .idMedico = oRsTmpHBT1.Fields!idMedico
                .IdProgramacion = oRsTmpHBT1.Fields!IdProgramacion
                .idServicio = IIf(IsNull(oRsTmpHBT1.Fields!idServicio), 0, oRsTmpHBT1.Fields!idServicio)
                .IdTipoProgramacion = IIf(IsNull(oRsTmpHBT1.Fields!IdTipoProgramacion), 0, oRsTmpHBT1.Fields!IdTipoProgramacion)
                .IdTipoServicio = oRsTmpHBT1.Fields!IdTipoServicio
                .IdTurno = IIf(IsNull(oRsTmpHBT1.Fields!IdTurno), 0, oRsTmpHBT1.Fields!IdTurno)
                .IdUsuarioAuditoria = lnIdUsuario
            End With
            If lnUltimoId < oDOProgramacionMedica.IdProgramacion Then
                If Not InsertarDebbProgramacionMedicaAgregar(oDOProgramacionMedica) Then
                      GoTo Terminar
                End If
            ElseIf Me.chkProgramacion.Value = 1 Then
                If Not oProgramacionMedica.Modificar(oDOProgramacionMedica) Then
                     ms_MensajeError = oProgramacionMedica.MensajeError: GoTo Terminar
                End If
            End If
            '
            oRsTmpHBT1.MoveNext
          Loop
       End If
       oRsTmpHBT1.Close
       '
       oConexion.CommitTrans
       Me.MousePointer = 1
       Unload Me
    End If
    Exit Sub
            
Terminar:
    oConexion.RollbackTrans
    MsgBox ms_MensajeError
    Me.MousePointer = 1
    Resume
    
End Sub

Private Sub Command10_Click()
    If MsgBox("Esta seguro de ELIMINAR", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        On Error GoTo ErrElim
        oRsFacturacionServicioFinanciamientos.Delete
        oRsFacturacionServicioFinanciamientos.Update
    End If
ErrElim:

End Sub

Private Sub Command11_Click()
    If MsgBox("Esta seguro de ELIMINAR", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        On Error GoTo ErrElim
        oRsFactOrdenServicioPagos.Delete
        oRsFactOrdenServicioPagos.Update
    End If
ErrElim:


End Sub

Private Sub Command12_Click()
    If MsgBox("Esta seguro de ELIMINAR", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        On Error GoTo ErrElim
        oRsFacturacionServicioPagos.Delete
        oRsFacturacionServicioPagos.Update
    End If
ErrElim:

End Sub

Private Sub Command13_Click()
    oRsFarmMovimientoVentas.AddNew
End Sub

Private Sub Command14_Click()
     oRsFacturacionBienesFinanciamiento.AddNew
End Sub

Private Sub Command15_Click()
   oRsFactOrdenesBienes.AddNew
End Sub

Private Sub Command16_Click()
    oRsFacturacionBienesPagos.AddNew
End Sub

Private Sub Command17_Click()
    oRsCajaComprobantePago.AddNew
End Sub

Private Sub Command18_Click()
    If MsgBox("Esta seguro de ELIMINAR", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        On Error GoTo ErrElim
        oRsFarmMovimientoVentas.Delete
        oRsFarmMovimientoVentas.Update
    End If
ErrElim:

End Sub

Private Sub Command19_Click()
    If MsgBox("Esta seguro de ELIMINAR", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        On Error GoTo ErrElim
        oRsFacturacionBienesFinanciamiento.Delete
        oRsFacturacionBienesFinanciamiento.Update
    End If
ErrElim:

End Sub

Private Sub Command2_Click()
    Dim oRs1 As New Recordset
    Dim oRs2 As New Recordset
    Dim oRs3 As New Recordset
    Dim oRs4 As New Recordset
    'INICIO-Actualiza Precio Venta Farmacia para los nuevos Tarifarios,
    '       en base a la tarifa idTipoFinanciamiento=1
    oRs4.Open "select * from FactCatalogoBienesInsumosHosp where idTipoFinanciamiento=1 and activo=1", sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
    oRs1.Open "select * from TiposFinanciamiento where SeIngresPrecios=1 and idTipoFinanciamiento>0", sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
    oRs4.MoveFirst
    Do While Not oRs4.EOF
       oRs1.MoveFirst
       Do While Not oRs1.EOF
            lcSql = "select * from FactCatalogoBienesInsumosHosp where idProducto=" & oRs4.Fields!idProducto & " and idTipoFinanciamiento=" & oRs1.Fields!IdTipoFinanciamiento
            oRs2.Open lcSql, sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
            If oRs2.RecordCount > 0 Then
               oRs2.Fields!PrecioUnitario = oRs4.Fields!PrecioUnitario
               oRs2.Update
            Else
               oRs2.AddNew
               oRs2.Fields!idProducto = oRs4.Fields!idProducto
               oRs2.Fields!IdTipoFinanciamiento = oRs1.Fields!IdTipoFinanciamiento
               oRs2.Fields!Activo = 1
               oRs2.Fields!PrecioUnitario = oRs4.Fields!PrecioUnitario
               oRs2.Update
            End If
            oRs2.Close
            oRs1.MoveNext
       Loop
       oRs4.MoveNext
   Loop
   Unload Me
End Sub



Private Sub Command20_Click()
    If MsgBox("Esta seguro de ELIMINAR", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        On Error GoTo ErrElim
        oRsFactOrdenesBienes.Delete
        oRsFactOrdenesBienes.Update
    End If
ErrElim:

End Sub

Private Sub Command21_Click()
    If MsgBox("Esta seguro de ELIMINAR", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        On Error GoTo ErrElim
        oRsFacturacionBienesPagos.Delete
        oRsFacturacionBienesPagos.Update
    End If
ErrElim:

End Sub

Private Sub Command22_Click()
    If MsgBox("Esta seguro de ELIMINAR", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        On Error GoTo ErrElim
        oRsCajaComprobantePagoS.Delete
        oRsCajaComprobantePagoS.Update
    End If
ErrElim:

End Sub

Private Sub Command23_Click()
    oRsFactOrdenServicio.AddNew
End Sub

Private Sub Command24_Click()
    oRsFacturacionServicioFinanciamientos.AddNew
End Sub

Private Sub Command25_Click()
    oRsFactOrdenServicioPagos.AddNew
End Sub

Private Sub Command26_Click()
    oRsFacturacionServicioPagos.AddNew
End Sub

Private Sub Command27_Click()
    oRsCajaComprobantePagoS.AddNew
End Sub

Private Sub Command28_Click()
   Dim oRsTmp0 As New Recordset
   Dim oRsTmp1 As New Recordset
   oRsTmp0.Open "select * from FarmMovimientoVentas where idfuenteFinanciamiento=14", sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
   If oRsTmp0.RecordCount > 0 Then
      oRsTmp0.MoveFirst
      Do While Not oRsTmp0.EOF
         lcSql = "update FarmMovimiento set idTipoConcepto=23 where movNumero='" & oRsTmp0.Fields!movNumero & "' and movTipo='" & oRsTmp0.Fields!MovTipo & "'"
         oRsTmp1.Open lcSql, sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
         oRsTmp0.MoveNext
      Loop
   End If
   oRsTmp0.Close
   oRsTmp1.Open "update FuentesFinanciamiento set idTipoConceptoFarmacia=23 where idFuenteFinanciamiento=14", sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
   Unload Me
End Sub

Sub GeneraOactualizaPuntoDeCargaPorCadaCPT(lnIdPuntoCarga As Long, lnIdServicioPuntoCarga As Long)
   Dim lcFiltraCPT As String
   Dim oRsFiltraCPT As New Recordset
   Dim oRsTmp1 As New Recordset
   Dim oRsTmp2 As New Recordset
   Dim DServicio As String
   'Genera nuevo Servicio
   If lnIdServicioPuntoCarga = 0 And (lnIdPuntoCarga = 32 Or lnIdPuntoCarga = 2 Or lnIdPuntoCarga = 31 Or lnIdPuntoCarga = 33 Or lnIdPuntoCarga = 34 Or lnIdPuntoCarga = 35 Or lnIdPuntoCarga = 36 Or lnIdPuntoCarga = 37) Then
        If lnIdPuntoCarga = 32 Then
           DServicio = "Anatomía Patológica"
        End If
        If lnIdPuntoCarga = 2 Then
           DServicio = "Patología Clínica"
        End If
        If lnIdPuntoCarga = 31 Then
           DServicio = "Citología"
        End If
        If lnIdPuntoCarga = 33 Then
           DServicio = "Microbiología"
        End If
        If lnIdPuntoCarga = 34 Then
           DServicio = "Hematología"
        End If
        If lnIdPuntoCarga = 35 Then
           DServicio = "Inmunoserología"
        End If
        If lnIdPuntoCarga = 36 Then
           DServicio = "Urianálisis y Parasitología"
        End If
        If lnIdPuntoCarga = 37 Then
           DServicio = "Bioquímica"
        End If
        lcSql = "select * from Servicios where nombre='" & DServicio & "'"
        oRsFiltraCPT.Open lcSql, sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
        If oRsFiltraCPT.RecordCount = 0 Then
           oRsFiltraCPT.AddNew
           oRsFiltraCPT.Fields!Nombre = DServicio
           oRsFiltraCPT.Fields!IdEspecialidad = 59
           oRsFiltraCPT.Fields!IdTipoServicio = 5
           oRsFiltraCPT.Fields!codigo = Trim(Str(oRsUltCodigo))
           oRsFiltraCPT.Fields!solotipoSexo = 0
           oRsFiltraCPT.Fields!maximaEdad = 54750
           oRsFiltraCPT.Fields!idEstado = 1
           oRsFiltraCPT.Update
           oRsUltCodigo = oRsUltCodigo + 1
        End If
        lnIdServicioPuntoCarga = oRsFiltraCPT.Fields!idServicio
        oRsFiltraCPT.Close
   End If
   'Actualiza IdServicio en tabla 'FactPuntoCarga'
   lcSql = "update FactPuntosCarga set idServicio=" & lnIdServicioPuntoCarga & " where idPuntoCarga=" & lnIdPuntoCarga
   oRsTmp1.Open lcSql, sighcomun.CadenaConexionShape, adOpenKeyset, adLockOptimistic
   '
   If lnIdServicioPuntoCarga > 0 And lnIdPuntoCarga > 0 Then
        lcFiltraCPT = "SELECT      dbo.FactCatalogoServicios.IdProducto, dbo.FactCatalogoServicios.Codigo, dbo.FactCatalogoServicios.Nombre, " & _
                 "                      dbo.FactCatalogoServiciosHosp.PrecioUnitario, dbo.FactCatalogoServiciosHosp.Activo, dbo.FactCatalogoServiciosPtos.idPuntoCarga," & _
                 "                      dbo.FactCatalogoServiciosHosp.SeUsaSinPrecio, dbo.FactCatalogoServicios.Nombre AS NombreProducto" & _
                 " FROM         dbo.FactCatalogoServicios RIGHT OUTER JOIN" & _
                 "                      dbo.FactCatalogoServiciosPtos ON dbo.FactCatalogoServicios.IdProducto = dbo.FactCatalogoServiciosPtos.idProducto RIGHT OUTER JOIN" & _
                 "                      dbo.FactCatalogoServiciosHosp ON dbo.FactCatalogoServicios.IdProducto = dbo.FactCatalogoServiciosHosp.IdProducto" & _
                 " WHERE     (dbo.FactCatalogoServiciosPtos.idPuntoCarga = " & lnIdPuntoCarga & ") AND (dbo.FactCatalogoServiciosHosp.IdTipoFinanciamiento = 1) AND" & _
                 "                      (dbo.FactCatalogoServicios.EsCPT = 1)" & _
                 " ORDER BY dbo.FactCatalogoServicios.Nombre"
         oRsTmp1.Open lcFiltraCPT, sighcomun.CadenaConexionShape, adOpenKeyset, adLockOptimistic
         If oRsTmp1.RecordCount > 0 Then
            Do While Not oRsTmp1.EOF
               lcSql = "select * from FactCatalogoServiciosPtos where idPuntoCarga=" & lnIdPuntoCarga & " and idProducto=" & oRsTmp1.Fields!idProducto
               oRsTmp2.Open lcSql, sighcomun.CadenaConexionShape, adOpenKeyset, adLockOptimistic
               If oRsTmp2.RecordCount > 0 Then
                  oRsTmp2.Fields!EsPreVenta = 1
                  oRsTmp2.Update
               Else
                  oRsTmp2.AddNew
                  oRsTmp2.Fields!idPuntoCarga = lnIdPuntoCarga
                  oRsTmp2.Fields!idProducto = oRsTmp1.Fields!idProducto
                  oRsTmp2.Fields!EsPreVenta = 1
                  oRsTmp2.Update
               End If
               oRsTmp2.Close
               oRsTmp1.MoveNext
            Loop
         End If
         oRsTmp1.Close
    End If
End Sub

Private Sub Command29_Click()
   Dim oRsFiltraCPT As New Recordset
   Dim lnIdServicio As Long, lnIdPuntoCarga As Long, lnIdServicioPuntoCarga As Long
   oRsUltCodigo = 999991
   'Agregar Servicio de "Estadistica" y actualizar en tabla "parametros"
   lcSql = "select * from Servicios where nombre='Estadística'"
   oRsFiltraCPT.Open lcSql, sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
   If oRsFiltraCPT.RecordCount = 0 Then
      oRsFiltraCPT.AddNew
      oRsFiltraCPT.Fields!Nombre = "Estadística"
      oRsFiltraCPT.Fields!IdEspecialidad = 93
      oRsFiltraCPT.Fields!IdTipoServicio = 1
      oRsFiltraCPT.Fields!codigo = Trim(Str(oRsUltCodigo))
      oRsFiltraCPT.Fields!solotipoSexo = 3
      oRsFiltraCPT.Fields!maximaEdad = 54750
      oRsFiltraCPT.Fields!idEstado = 1
      oRsFiltraCPT.Update
   End If
   lnIdServicio = oRsFiltraCPT.Fields!idServicio
   oRsFiltraCPT.Close
   lcSql = "update parametros set valorTexto='" & Trim(Str(lnIdServicio)) & "' where idParametro=256"
   oRsFiltraCPT.Open lcSql, sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
   oRsUltCodigo = oRsUltCodigo + 1
   'Rx
   lnIdPuntoCarga = 21
   lnIdServicioPuntoCarga = 23
   GeneraOactualizaPuntoDeCargaPorCadaCPT lnIdPuntoCarga, lnIdServicioPuntoCarga
   'Tomografia
   lnIdPuntoCarga = 22
   lnIdServicioPuntoCarga = 22
   GeneraOactualizaPuntoDeCargaPorCadaCPT lnIdPuntoCarga, lnIdServicioPuntoCarga
   'Ecog General
   lnIdPuntoCarga = 20
   lnIdServicioPuntoCarga = 24
   GeneraOactualizaPuntoDeCargaPorCadaCPT lnIdPuntoCarga, lnIdServicioPuntoCarga
   'Ecog.Obstetrica
   lnIdPuntoCarga = 23
   lnIdServicioPuntoCarga = 95
   GeneraOactualizaPuntoDeCargaPorCadaCPT lnIdPuntoCarga, lnIdServicioPuntoCarga
   'Banco Sangre
   lnIdPuntoCarga = 38
   lnIdServicioPuntoCarga = 19
   GeneraOactualizaPuntoDeCargaPorCadaCPT lnIdPuntoCarga, lnIdServicioPuntoCarga
   'Anat.Patologica
   lnIdPuntoCarga = 32
   lnIdServicioPuntoCarga = 0
   GeneraOactualizaPuntoDeCargaPorCadaCPT lnIdPuntoCarga, lnIdServicioPuntoCarga
   'Patologia clinica
   lnIdPuntoCarga = 2
   lnIdServicioPuntoCarga = 0
   GeneraOactualizaPuntoDeCargaPorCadaCPT lnIdPuntoCarga, lnIdServicioPuntoCarga
   'Citologia
   lnIdPuntoCarga = 31
   lnIdServicioPuntoCarga = 0
   GeneraOactualizaPuntoDeCargaPorCadaCPT lnIdPuntoCarga, lnIdServicioPuntoCarga
   'Microbiologia
   lnIdPuntoCarga = 33
   lnIdServicioPuntoCarga = 0
   GeneraOactualizaPuntoDeCargaPorCadaCPT lnIdPuntoCarga, lnIdServicioPuntoCarga
   'Hematologia
   lnIdPuntoCarga = 34
   lnIdServicioPuntoCarga = 0
   GeneraOactualizaPuntoDeCargaPorCadaCPT lnIdPuntoCarga, lnIdServicioPuntoCarga
   'InmunoSerologia
   lnIdPuntoCarga = 35
   lnIdServicioPuntoCarga = 0
   GeneraOactualizaPuntoDeCargaPorCadaCPT lnIdPuntoCarga, lnIdServicioPuntoCarga
   'Urianalisis y parasitologia
   lnIdPuntoCarga = 36
   lnIdServicioPuntoCarga = 0
   GeneraOactualizaPuntoDeCargaPorCadaCPT lnIdPuntoCarga, lnIdServicioPuntoCarga
   'Bioquimica
   lnIdPuntoCarga = 37
   lnIdServicioPuntoCarga = 0
   GeneraOactualizaPuntoDeCargaPorCadaCPT lnIdPuntoCarga, lnIdServicioPuntoCarga
   '
   Unload Me
ErrorC7:
End Sub

Private Sub Command3_Click()
    Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
    txtGalenHos.Text = Trim(Str(mo_ReglasFacturacion.RetornaConsumoPacienteServiciosConSeguroPorNroCuenta(Val(txtGalenHos.Text))))
End Sub

Private Sub Command30_Click()
    Dim oRsTmp As New Recordset
    Dim oRsTmp1 As New Recordset
    Dim lnIdPaciente As Long
    'Actualiza datos personales de pacientes en Tablas de Laboratorio e Imagenes
    oRsTmp.Open "Select * from LabMovimientoLaboratorio", sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          If oRsTmp.Fields!idCuentaAtencion > 0 Then
                lcSql = "SELECT     dbo.LabMovimientoLaboratorio.IdMovimiento, dbo.LabMovimientoLaboratorio.IdCuentaAtencion, dbo.Atenciones.IdPaciente" & _
                        " FROM         dbo.LabMovimientoLaboratorio LEFT OUTER JOIN" & _
                        "           dbo.Atenciones ON dbo.LabMovimientoLaboratorio.IdCuentaAtencion = dbo.Atenciones.IdCuentaAtencion" & _
                        " Where dbo.LabMovimientoLaboratorio.IdCuentaAtencion=" & oRsTmp.Fields!idCuentaAtencion
          Else
                lcSql = "SELECT     IdOrden, IdPaciente" & _
                        " From dbo.FactOrdenServicio" & _
                        " Where IdOrden = " & oRsTmp.Fields!IdOrden
          End If
          oRsTmp1.Open lcSql, sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
          lnIdPaciente = 0
          If oRsTmp1.RecordCount > 0 Then
             If oRsTmp1.Fields!idPaciente > 0 Then
                lnIdPaciente = oRsTmp1.Fields!idPaciente
             End If
          End If
          oRsTmp1.Close
          If lnIdPaciente > 0 Then
                oRsTmp1.Open "Select * from Pacientes where idPaciente=" & lnIdPaciente, sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
                If oRsTmp1.RecordCount > 0 Then
                   oRsTmp.Fields!Paciente = Left(Trim(oRsTmp1.Fields!apellidoPaterno) & " " & Trim(oRsTmp1.Fields!apellidoMaterno) & " " & Trim(oRsTmp1.Fields!PrimerNombre), 100)
                   oRsTmp.Fields!idTipoSexo = oRsTmp1.Fields!idTipoSexo
                   oRsTmp.Fields!FechaNacimiento = oRsTmp1.Fields!FechaNacimiento
                   oRsTmp.Update
                End If
                oRsTmp1.Close
          End If
          oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    oRsTmp.Open "Select * from ImagMovimientoImagenes", sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          If oRsTmp.Fields!idCuentaAtencion > 0 Then
                lcSql = "SELECT     dbo.ImagMovimientoImagenes.IdMovimiento, dbo.Atenciones.IdPaciente, dbo.ImagMovimientoImagenes.IdCuentaAtencion" & _
                        " FROM         dbo.ImagMovimientoImagenes LEFT OUTER JOIN" & _
                        "        dbo.Atenciones ON dbo.ImagMovimientoImagenes.IdCuentaAtencion = dbo.Atenciones.IdCuentaAtencion" & _
                        " where dbo.ImagMovimientoImagenes.IdCuentaAtencion=" & oRsTmp.Fields!idCuentaAtencion
          Else
                lcSql = "SELECT     dbo.ImagMovimientoImagenes.IdMovimiento, dbo.ImagMovimientoImagenes.IdComprobantePago, dbo.FactOrdenServicio.IdPaciente, " & _
                        "             dbo.ImagMovimientoImagenes.IdOrden" & _
                        " FROM         dbo.ImagMovimientoImagenes LEFT OUTER JOIN" & _
                        "     dbo.FactOrdenServicio ON dbo.ImagMovimientoImagenes.IdOrden = dbo.FactOrdenServicio.IdOrden" & _
                        " where dbo.ImagMovimientoImagenes.IdOrden=" & oRsTmp.Fields!IdOrden
          End If
          oRsTmp1.Open lcSql, sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
          lnIdPaciente = 0
          If oRsTmp1.RecordCount > 0 Then
             If oRsTmp1.Fields!idPaciente > 0 Then
                lnIdPaciente = oRsTmp1.Fields!idPaciente
             End If
          End If
          oRsTmp1.Close
          If lnIdPaciente > 0 Then
                oRsTmp1.Open "Select * from Pacientes where idPaciente=" & lnIdPaciente, sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
                If oRsTmp1.RecordCount > 0 Then
                   oRsTmp.Fields!Paciente = Left(Trim(oRsTmp1.Fields!apellidoPaterno) & " " & Trim(oRsTmp1.Fields!apellidoMaterno) & " " & Trim(oRsTmp1.Fields!PrimerNombre), 100)
                   oRsTmp.Fields!idTipoSexo = oRsTmp1.Fields!idTipoSexo
                   oRsTmp.Fields!FechaNacimiento = oRsTmp1.Fields!FechaNacimiento
                   oRsTmp.Update
                End If
                oRsTmp1.Close
          End If
          oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    Unload Me
End Sub

Private Sub Command31_Click()
    Dim oRsTmp As New Recordset
    Dim oRsTmp1 As New Recordset
    Dim lnIdPaciente As Long
    'busca N° Historia para Pacientes con Historia=NULL
    oRsTmp.Open "Select * from Pacientes where idTipoNumeracion<4 and NroHistoriaClinica IS NULL", sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          lcSql = "select * from HistoriasClinicas where idPaciente=" & oRsTmp.Fields!idPaciente
          oRsTmp1.Open lcSql, sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
          If oRsTmp1.RecordCount > 0 Then
             oRsTmp.Fields!NroHistoriaClinica = oRsTmp1.Fields!NroHistoriaClinica
             oRsTmp.Update
          End If
          oRsTmp1.Close
          oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    'Asigna atenciones.RecienNacido=1 para los que tienen hasta 1 mes de nacido
    oRsTmp.Open "Select * from Atenciones", sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          oRsTmp1.Open "select FechaNacimiento from Pacientes where idPaciente=" & oRsTmp.Fields!idPaciente, sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
          If oRsTmp1.RecordCount > 0 Then
             If sighcomun.CalculaSiEsRecienNacido(oRsTmp1.Fields!FechaNacimiento, CDate(oRsTmp.Fields!FechaIngreso & " " & oRsTmp.Fields!HoraIngreso)) = 1 Then
                oRsTmp.Fields!RecienNacido = 1
             Else
                oRsTmp.Fields!RecienNacido = 0
             End If
             oRsTmp.Update
          End If
          oRsTmp1.Close
          oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    '
    Unload Me

End Sub

Private Sub Command32_Click()
    Dim oRsTmp As New Recordset
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim lcSerie As String, lcDcto As String, ldFecha As Date
    Dim lbProceso As Boolean
    lcSql = "SELECT      dbo.CajaComprobantesPago.NroSerie, dbo.CajaComprobantesPago.NroDocumento, dbo.FactOrdenesBienes.idPuntoCarga, " & _
            "                      dbo.FacturacionBienesPagos.IdOrden, dbo.FacturacionBienesPagos.IdProducto, dbo.FacturacionBienesPagos.CantidadPagar," & _
            "                      dbo.FacturacionBienesPagos.PrecioVenta, dbo.FacturacionBienesPagos.TotalPagar, dbo.FactOrdenesBienes.idOrden," & _
            "                      dbo.FactOrdenesBienes.idCuentaAtencion, dbo.FactOrdenesBienes.idPreventa, dbo.CajaComprobantesPago.IdTipoOrden," & _
            "                      dbo.CajaComprobantesPago.Total , dbo.CajaComprobantesPago.IdEstadoComprobante, dbo.CajaComprobantesPago.FechaCobranza" & _
            " FROM         dbo.FactOrdenesBienes LEFT OUTER JOIN" & _
            "                      dbo.CajaComprobantesPago ON dbo.FactOrdenesBienes.idComprobantePago = dbo.CajaComprobantesPago.IdComprobantePago LEFT OUTER JOIN" & _
            "                      dbo.FacturacionBienesPagos ON dbo.FactOrdenesBienes.idOrden = dbo.FacturacionBienesPagos.IdOrden" & _
            " Where (dbo.CajaComprobantesPago.IdEstadoComprobante = 9)" & _
            " ORDER BY dbo.CajaComprobantesPago.NroSerie, dbo.CajaComprobantesPago.NroDocumento"
     oRsTmp.Open lcSql, sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
     If oRsTmp.RecordCount > 0 Then
        With wrs_Gal
            .Fields.Append "Serie", adVarChar, 5, adFldIsNullable
            .Fields.Append "Documento", adVarChar, 20, adFldIsNullable
            .Fields.Append "Fecha", adDate
            .LockType = adLockOptimistic
            .Open
        End With
        oRsTmp.MoveFirst
        Do While Not oRsTmp.EOF
           lcSerie = oRsTmp.Fields!nroSerie
           lcDcto = oRsTmp.Fields!NroDocumento
           ldFecha = oRsTmp.Fields!FechaCobranza
           lbProceso = False
           Do While Not oRsTmp.EOF And lcSerie = oRsTmp.Fields!nroSerie And lcDcto = oRsTmp.Fields!NroDocumento
              If oRsTmp.Fields!idPreVenta > 0 Then
                 lcSql = "select * from FarmPreventa where idPreventa=" & oRsTmp.Fields!idPreVenta
                 oRsTmp1.Open lcSql, sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
                 If oRsTmp1.RecordCount > 0 And oRsTmp1.Fields!idEstadoPreventa = 1 Then
                    lbProceso = True
                    lcSql = "update FactORdenesBienes set idComprobantePago=null where idOrden=" & oRsTmp.Fields!IdOrden
                    oRsTmp2.Open lcSql, sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
                 End If
                 oRsTmp1.Close
              End If
              oRsTmp.MoveNext
              If oRsTmp.EOF Then
                 Exit Do
              End If
           Loop
           If lbProceso = True Then
              wrs_Gal.AddNew
              wrs_Gal.Fields!Serie = lcSerie
              wrs_Gal.Fields!Documento = lcDcto
              wrs_Gal.Fields!Fecha = ldFecha
              wrs_Gal.Update
           End If
        Loop
     End If
     oRsTmp.Close
     Set grdGalenHos.DataSource = wrs_Gal
     MsgBox "Comprueba que la lista de BOLETAS estan bien reparadas"
End Sub

Private Sub Command4_Click()
  wrs_Gal.AddNew
'  wrs_Gal.Update
End Sub


Private Sub Command1_Click()
   If MsgBox("Esta seguro?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        wrs_Gal.Delete
        wrs_Gal.Update
        wrs_Gal.Requery
    End If
End Sub

Private Sub Command5_Click()
      Dim oCrypKey As New CrypKey.Util
      MsgBox oCrypKey.DecryptString(txtGalenHos.Text)
End Sub

Private Sub Command6_Click()
    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    txtGalenHos.Text = Trim(Str(mo_ReglasFarmacia.RetornaConsumoPacienteFarmaciaConSeguroPorNroCuenta(Val(txtGalenHos.Text))))

End Sub

Private Sub Command7_Click()
    Dim lcNroHistoriaNew As String
    lcNroHistoriaNew = InputBox("ingrese N° Historia NUEVA: ")
    If Val(txtGalenHos.Text) = 0 Then
       MsgBox "ingrese el N° Historia ACTUAL en texto SQL"
       Exit Sub
    End If
    If Val(lcNroHistoriaNew) = 0 Then
       MsgBox "ingrese el N° Historia NUEVA"
       Exit Sub
    End If
    On Error GoTo ErrorC7
    Dim oRsTmp As New Recordset
    Dim oConexion As New Connection
    oConexion.Open sighcomun.CadenaConexion
    oConexion.BeginTrans
    oRsTmp.Open "update historiasClinicas set NroHistoriaClinica=" & lcNroHistoriaNew & " where nroHistoriaClinica=" & txtGalenHos.Text, oConexion, adOpenKeyset, adLockOptimistic
    oRsTmp.Open "update pacientes set NroHistoriaClinica=" & lcNroHistoriaNew & " where nroHistoriaClinica=" & txtGalenHos.Text, oConexion, adOpenKeyset, adLockOptimistic
    oConexion.CommitTrans
    Exit Sub
ErrorC7:
    oConexion.RollbackTrans
    MsgBox Err.Description
End Sub

Private Sub Command8_Click()
    If MsgBox("Esta seguro de ELIMINAR", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        On Error GoTo ErrElim
        oRsFactOrdenServicio.Delete
        oRsFactOrdenServicio.Update
    End If
ErrElim:
End Sub

Private Sub Command9_Click()
    If MsgBox("Esta seguro de ELIMINAR", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        On Error GoTo ErrElim
        oRsCajaComprobantePagoS.Delete
        oRsCajaComprobantePagoS.Update
    End If
ErrElim:

End Sub

Private Sub Form_Load()
    Set grdFarmMovimientoVentas.DataSource = oRsFarmMovimientoVentas
    Set grdCajaComprobantesPago.DataSource = oRsCajaComprobantePago
    Set grdFacturacionBienesFinanciamiento.DataSource = oRsFacturacionBienesFinanciamiento
    Set grdFactOrdenesBienes.DataSource = oRsFactOrdenesBienes
    Set grdFacturacionBienesPagos.DataSource = oRsFacturacionBienesPagos
    
    Set grdFactOrdenServicio.DataSource = oRsFactOrdenServicio
    Set grdCajaComprobantesPagoS.DataSource = oRsCajaComprobantePagoS
    Set grdFacturacionServicioFinanciamientos.DataSource = oRsFacturacionServicioFinanciamientos
    Set grdFactOrdenServicioPagos.DataSource = oRsFactOrdenServicioPagos
    Set grdFacturacionServicioPagos.DataSource = oRsFacturacionServicioPagos
    '
    cmbConsideraciones.AddItem "Consideraciones:"
    cmbConsideraciones.AddItem ""
    cmbConsideraciones.AddItem "1-Se elimina las Atenciones de esas Fechas,"
    cmbConsideraciones.AddItem "  se agrega las Atenciones de esas Fechas. "
    cmbConsideraciones.AddItem "2-Deberá tener actualizado las tablas:     "
    cmbConsideraciones.AddItem "  empleados,Medicos,Especialidades,MedicosEspecialidad,"
    cmbConsideraciones.AddItem "  EstablecimientosNoMinsa,Servicios,"
    cmbConsideraciones.AddItem "  camas,.."
    cmbConsideraciones.AddItem "  FactCatalogoBienesInsumos,FactCatalogoServicios,"
    cmbConsideraciones.AddItem "  FuentesFinanciamiento, FuentesFinanciamientoTarifas,"
    cmbConsideraciones.AddItem "  Turnos"
    cmbConsideraciones.AddItem ""
    cmbConsideraciones.AddItem "3-Quitar Autogenerado de: Atenciones, atencionesEmergencia,,"
    cmbConsideraciones.AddItem "  AtencionesDiagnosticos , atencionesEstanciaHospitalaria"
    cmbConsideraciones.AddItem "  atencionesNacimientos, citas, camas,  EstablecimientosNoMinsa "
    cmbConsideraciones.AddItem "  empleados,    "
    cmbConsideraciones.AddItem "  FacturacionCuentasAtencion, "
    cmbConsideraciones.AddItem "  HistoriasSolicitadas, MovimientosHistoriaClinica, "
    cmbConsideraciones.AddItem "  Medicos, MedicosEspecialidad,"
    cmbConsideraciones.AddItem "  ProgramacionMedica, Pacientes,"
    cmbConsideraciones.AddItem ""
    cmbConsideraciones.AddItem "4-Cuando ya se termine totalmente de migrar:"
    cmbConsideraciones.AddItem "  * ejecutar 'actualiza AUTOGENERADOS.SQL'"
    cmbConsideraciones.AddItem "  * eliminar Proced.Almac. que empiezen con DEBB..."
    cmbConsideraciones.AddItem "  * actualizar correlativos en: GeneradorNroHistoriaClinica"
    
End Sub

Sub LimpiaGrid()
    Set grdFarmMovimientoVentas.DataSource = Nothing
    Set grdCajaComprobantesPago.DataSource = Nothing
    Set grdFacturacionBienesFinanciamiento.DataSource = Nothing
    Set grdFactOrdenesBienes.DataSource = Nothing
    Set grdFacturacionBienesPagos.DataSource = Nothing
End Sub

Sub LimpiaGridS()
    Set grdFactOrdenServicio.DataSource = Nothing
    Set grdCajaComprobantesPagoS.DataSource = Nothing
    Set grdFacturacionServicioFinanciamientos.DataSource = Nothing
    Set grdFactOrdenServicioPagos.DataSource = Nothing
    Set grdFacturacionServicioPagos.DataSource = Nothing

End Sub

Private Sub grdFactOrdenesBienes_DblClick()
        On Error GoTo ErrFOB
        Dim lnLinea As Integer
        lnLinea = 1
        lcSql = "select * from FacturacionBienesPagos where idOrden=" & oRsFactOrdenesBienes.Fields!IdOrden
        oRsFacturacionBienesPagos.Open lcSql, sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
        Set grdFacturacionBienesPagos.DataSource = oRsFacturacionBienesPagos
        lnLinea = 2
       If oRsFactOrdenesBienes.Fields!IdComprobantePago > 0 Then
            lcSql = "select * from CajaComprobantesPago where idComprobantePago=" & oRsFactOrdenesBienes.Fields!IdComprobantePago
            oRsCajaComprobantePago.Open lcSql, sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
            Set grdCajaComprobantesPago.DataSource = oRsCajaComprobantePago
       Else
            Set grdCajaComprobantesPago.DataSource = Nothing
       End If
    Exit Sub
ErrFOB:
   If Err.Number = 3705 Then
      Select Case lnLinea
      Case 1
           oRsFacturacionBienesPagos.Close
      Case 2
           oRsCajaComprobantePago.Close
      End Select
      Resume
   End If
End Sub

Private Sub grdFactOrdenServicio_DblClick()
       On Error GoTo ErrFMV
       Dim lnLinea As Integer
       lnLinea = 1
       lcSql = "select * from FactOrdenServicioPagos where idOrden=" & oRsFactOrdenServicio.Fields!IdOrden
       oRsFactOrdenServicioPagos.Open lcSql, sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
       Set grdFactOrdenServicioPagos.DataSource = oRsFactOrdenServicioPagos
       lnLinea = 2
       lcSql = "select * from FacturacionServicioFinanciamientos where idOrden=" & oRsFactOrdenServicio.Fields!IdOrden
       oRsFacturacionServicioFinanciamientos.Open lcSql, sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
       Set grdFacturacionServicioFinanciamientos.DataSource = oRsFacturacionServicioFinanciamientos
       
       Set grdCajaComprobantesPagoS.DataSource = Nothing
       Set grdFacturacionServicioPagos.DataSource = Nothing
       
    Exit Sub
ErrFMV:
   If Err.Number = 3705 Then
      Select Case lnLinea
      Case 1
          oRsFactOrdenServicioPagos.Close
      Case 2
          oRsFacturacionServicioFinanciamientos.Close
      End Select
      Resume
   End If
End Sub

Private Sub grdFactOrdenServicioPagos_DblClick()
       On Error GoTo ErrFOB
       Dim lnLinea As Integer
       lnLinea = 1
       lcSql = "select * from FacturacionServicioPagos where idOrdenPago=" & oRsFactOrdenServicioPagos.Fields!IdOrdenPago
       oRsFacturacionServicioPagos.Open lcSql, sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
       Set grdFacturacionServicioPagos.DataSource = oRsFacturacionServicioPagos
       lnLinea = 2
       If oRsFactOrdenServicioPagos.Fields!IdComprobantePago > 0 Then
            lcSql = "select * from CajaComprobantesPago where idComprobantePago=" & oRsFactOrdenServicioPagos.Fields!IdComprobantePago
            oRsCajaComprobantePagoS.Open lcSql, sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
            Set grdCajaComprobantesPagoS.DataSource = oRsCajaComprobantePagoS
       Else
            Set grdCajaComprobantesPagoS.DataSource = Nothing
       End If
    Exit Sub
ErrFOB:
   If Err.Number = 3705 Then
      Select Case lnLinea
      Case 1
           oRsFacturacionServicioPagos.Close
      Case 2
           oRsCajaComprobantePagoS.Close
      End Select
      Resume
   End If

End Sub

Private Sub grdFarmMovimientoVentas_DblClick()
       On Error GoTo ErrFMV
       Dim lnLinea As Integer
       lnLinea = 1
       lcSql = "select * from FactOrdenesBienes where movNumero='" & oRsFarmMovimientoVentas.Fields!movNumero & "' and movTipo='" & oRsFarmMovimientoVentas.Fields!MovTipo & "'"
       oRsFactOrdenesBienes.Open lcSql, sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
       Set grdFactOrdenesBienes.DataSource = oRsFactOrdenesBienes
       lnLinea = 2
       lcSql = "select * from FacturacionBienesFinanciamientos where movNumero='" & oRsFarmMovimientoVentas.Fields!movNumero & "' and movTipo='" & oRsFarmMovimientoVentas.Fields!MovTipo & "'"
       oRsFacturacionBienesFinanciamiento.Open lcSql, sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
       Set grdFacturacionBienesFinanciamiento.DataSource = oRsFacturacionBienesFinanciamiento
       
       Set grdCajaComprobantesPago.DataSource = Nothing
       Set grdFacturacionBienesPagos.DataSource = Nothing
    Exit Sub
ErrFMV:
   If Err.Number = 3705 Then
      Select Case lnLinea
      Case 1
          oRsFactOrdenesBienes.Close
      Case 2
          oRsFacturacionBienesFinanciamiento.Close
      End Select
      Resume
   End If
End Sub





Private Sub txtCuentaS_KeyPress(KeyAscii As Integer)
    On Error GoTo errCta
    If KeyAscii = 13 And Val(txtCuentaS.Text) > 0 Then
       LimpiaGridS
       Dim lnLinea As Integer
       lnLinea = 1
       lcSql = "select * from FactOrdenServicio where idCuentaAtencion=" & txtCuentaS.Text
       oRsFactOrdenServicio.Open lcSql, sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
       Set grdFactOrdenServicio.DataSource = oRsFactOrdenServicio
       lnLinea = 2
       lcSql = "select * from CajaComprobantesPago where IdTipoOrden=1 and idCuentaAtencion=" & txtCuentaS.Text
       oRsCajaComprobantePagoS.Open lcSql, sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
       Set grdCajaComprobantesPagoS.DataSource = oRsCajaComprobantePagoS
    End If
    Exit Sub
errCta:
   If Err.Number = 3705 Then
      Select Case lnLinea
      Case 1
           oRsFactOrdenServicio.Close
      Case 2
           oRsCajaComprobantePagoS.Close
      End Select
      Resume
   End If

End Sub

Private Sub txtGalenHos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       usarSelectGalenHos txtGalenHos.Text
    End If

End Sub

Sub usarSelectGalenHos(txt As String)
    Dim wrs_Prg As New ADODB.Recordset
    On Error GoTo eRRCarga2
    wrs_Gal.Open txt, sighcomun.CadenaConexionShape, adOpenKeyset, adLockOptimistic
    Set grdGalenHos.DataSource = wrs_Gal
    Exit Sub
eRRCarga2:
    If Err.Number = 3705 Then
       wrs_Gal.Close
       Resume
    End If
End Sub



Private Sub txtNroCuenta_KeyPress(KeyAscii As Integer)
    On Error GoTo errCta
    If KeyAscii = 13 And Val(txtNroCuenta.Text) > 0 Then
       LimpiaGrid
       Dim lnLinea As Integer
       lnLinea = 1
       Set grdFarmMovimientoVentas.DataSource = Nothing
       lcSql = "select * from FarmMovimientoVentas where idCuentaAtencion=" & txtNroCuenta.Text
       oRsFarmMovimientoVentas.Open lcSql, sighcomun.CadenaConexionShape, adOpenKeyset, adLockOptimistic
       Set grdFarmMovimientoVentas.DataSource = oRsFarmMovimientoVentas
       lnLinea = 2
       lcSql = "select * from CajaComprobantesPago where IdTipoOrden<>1 and idCuentaAtencion=" & txtCuentaS.Text
       oRsCajaComprobantePago.Open lcSql, sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
       Set grdCajaComprobantesPago.DataSource = oRsCajaComprobantePago
    End If
    Exit Sub
errCta:
   If Err.Number = 3705 Then
      Select Case lnLinea
      Case 1
           oRsFarmMovimientoVentas.Close
      Case 2
           oRsCajaComprobantePago.Close
      End Select
      Resume
   End If
End Sub




Function InsertarTmpPacientesAgregar(ByVal oTabla As doPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   InsertarTmpPacientesAgregar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_conexion
       .CommandText = "debbPacientesAgregar"
           Set oParameter = .CreateParameter("@IdPaisNacimiento", adInteger, adParamInput, 0, IIf(oTabla.IdPaisNacimiento = 0, Null, oTabla.IdPaisNacimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ApellidoMaterno", adVarChar, adParamInput, 20, IIf(oTabla.apellidoMaterno = "", Null, oTabla.apellidoMaterno)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@DireccionDomicilio", adVarChar, adParamInput, 50, IIf(oTabla.DireccionDomicilio = "", Null, oTabla.DireccionDomicilio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Observacion", adVarChar, adParamInput, 150, IIf(oTabla.Observacion = "", Null, oTabla.Observacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoNumeracion", adInteger, adParamInput, 0, IIf(oTabla.IdTipoNumeracion = 0, Null, oTabla.IdTipoNumeracion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaisProcedencia", adInteger, adParamInput, 0, IIf(oTabla.IdPaisProcedencia = 0, Null, oTabla.IdPaisProcedencia)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, oTabla.idPaciente): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ApellidoPaterno", adVarChar, adParamInput, 20, IIf(oTabla.apellidoPaterno = "", Null, oTabla.apellidoPaterno)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@PrimerNombre", adVarChar, adParamInput, 20, IIf(oTabla.PrimerNombre = "", Null, oTabla.PrimerNombre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@SegundoNombre", adVarChar, adParamInput, 20, IIf(oTabla.SegundoNombre = "", Null, oTabla.SegundoNombre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TercerNombre", adVarChar, adParamInput, 20, IIf(oTabla.TercerNombre = "", Null, oTabla.TercerNombre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaNacimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaNacimiento = 0, Null, oTabla.FechaNacimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 12, IIf(oTabla.NroDocumento = "", Null, oTabla.NroDocumento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Telefono", adVarChar, adParamInput, 10, IIf(oTabla.Telefono = "", Null, oTabla.Telefono)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Autogenerado", adVarChar, adParamInput, 20, IIf(oTabla.Autogenerado = "", Null, oTabla.Autogenerado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoSexo", adInteger, adParamInput, 4, IIf(oTabla.idTipoSexo = 0, Null, oTabla.idTipoSexo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProcedencia", adInteger, adParamInput, 4, IIf(oTabla.IdProcedencia = 0, Null, oTabla.IdProcedencia)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdGradoInstruccion", adInteger, adParamInput, 4, IIf(oTabla.IdGradoInstruccion = 0, Null, oTabla.IdGradoInstruccion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstadoCivil", adInteger, adParamInput, 4, IIf(oTabla.IdEstadoCivil = 0, Null, oTabla.IdEstadoCivil)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDocIdentidad", adInteger, adParamInput, 4, IIf(oTabla.IdDocIdentidad = 0, Null, oTabla.IdDocIdentidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoOcupacion", adInteger, adParamInput, 4, IIf(oTabla.idTipoOcupacion = 0, Null, oTabla.idTipoOcupacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCentroPobladoDomicilio", adInteger, adParamInput, 4, IIf(oTabla.IdCentroPobladoDomicilio = 0, Null, oTabla.IdCentroPobladoDomicilio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NombrePadre", adVarChar, adParamInput, 20, IIf(oTabla.NombrePadre = "", Null, oTabla.NombrePadre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NombreMadre", adVarChar, adParamInput, 20, IIf(oTabla.NombreMadre = "", Null, oTabla.NombreMadre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaisDomicilio", adInteger, adParamInput, 4, IIf(oTabla.IdPaisDomicilio = 0, Null, oTabla.IdPaisDomicilio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 4, IIf(oTabla.NroHistoriaClinica = 0, Null, oTabla.NroHistoriaClinica)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCentroPobladoNacimiento", adInteger, adParamInput, 0, IIf(oTabla.IdCentroPobladoNacimiento = 0, Null, oTabla.IdCentroPobladoNacimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCentroPobladoProcedencia", adInteger, adParamInput, 0, IIf(oTabla.IdCentroPobladoProcedencia = 0, Null, oTabla.IdCentroPobladoProcedencia)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDistritoProcedencia", adInteger, adParamInput, 0, IIf(oTabla.IdDistritoProcedencia = 0, Null, oTabla.IdDistritoProcedencia)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDistritoDomicilio", adInteger, adParamInput, 0, IIf(oTabla.IdDistritoDomicilio = 0, Null, oTabla.IdDistritoDomicilio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDistritoNacimiento", adInteger, adParamInput, 0, IIf(oTabla.IdDistritoNacimiento = 0, Null, oTabla.IdDistritoNacimiento)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
   End With
   InsertarTmpPacientesAgregar = True
 Exit Function
ManejadorDeError:
   MsgBox Err.Number & " " + Err.Description
Exit Function
End Function

Function InsertarDebbProgramacionMedicaAgregar(ByVal oTabla As DOProgramacionMedica) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   InsertarDebbProgramacionMedicaAgregar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_conexion
       .CommandText = "debbProgramacionMedicaAgregar"
           Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, 0, IIf(oTabla.IdEspecialidad = 0, Null, oTabla.IdEspecialidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, IIf(oTabla.IdTurno = 0, Null, oTabla.IdTurno)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Color", adInteger, adParamInput, 0, IIf(oTabla.Color = 0, Null, oTabla.Color)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, IIf(oTabla.idServicio = 0, Null, oTabla.idServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProgramacion", adInteger, adParamInput, 0, oTabla.IdProgramacion): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, IIf(oTabla.idMedico = 0, Null, oTabla.idMedico)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDepartamento", adInteger, adParamInput, 0, IIf(oTabla.IdDepartamento = 0, Null, oTabla.IdDepartamento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoServicio", adInteger, adParamInput, 0, IIf(oTabla.IdTipoServicio = 0, Null, oTabla.IdTipoServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.Fecha = 0, Null, oTabla.Fecha)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraInicio", adChar, adParamInput, 5, IIf(oTabla.HoraInicio = "", Null, oTabla.HoraInicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraFin", adChar, adParamInput, 5, IIf(oTabla.HoraFin = "", Null, oTabla.HoraFin)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 100, IIf(oTabla.Descripcion = "", Null, oTabla.Descripcion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoProgramacion", adInteger, adParamInput, 0, IIf(oTabla.IdTipoProgramacion = 0, Null, oTabla.IdTipoProgramacion)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
   End With
   InsertarDebbProgramacionMedicaAgregar = True
Exit Function
ManejadorDeError:
   MsgBox Err.Number & " " + Err.Description
Exit Function
End Function


Function InsertarDebbCuentaAtencion(ByVal oTabla As DOCuentaAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   InsertarDebbCuentaAtencion = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_conexion
       .CommandText = "debbFacturacionCuentasAtencionAgregar"
           Set oParameter = .CreateParameter("@TotalPorPagar", adCurrency, adParamInput, 0, IIf(oTabla.TotalPorPagar = 0, Null, oTabla.TotalPorPagar)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstado", adInteger, adParamInput, 0, IIf(oTabla.idEstado = 0, Null, oTabla.idEstado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TotalPagado", adCurrency, adParamInput, 0, IIf(oTabla.TotalPagado = 0, Null, oTabla.TotalPagado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TotalAsegurado", adCurrency, adParamInput, 0, IIf(oTabla.TotalAsegurado = 0, Null, oTabla.TotalAsegurado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TotalExonerado", adCurrency, adParamInput, 0, IIf(oTabla.TotalExonerado = 0, Null, oTabla.TotalExonerado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraCierre", adChar, adParamInput, 5, IIf(oTabla.HoraCierre = "", Null, oTabla.HoraCierre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaCierre", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCierre = 0, Null, oTabla.FechaCierre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraApertura", adChar, adParamInput, 5, IIf(oTabla.HoraApertura = "", Null, oTabla.HoraApertura)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaApertura", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaApertura = 0, Null, oTabla.FechaApertura)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.idPaciente = 0, Null, oTabla.idPaciente)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oTabla.idCuentaAtencion): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
   End With
   InsertarDebbCuentaAtencion = True
Exit Function
ManejadorDeError:
   lnErrCA = Err.Number
   If lnErrCA <> -2147217873 Then
      MsgBox Err.Number & " " + Err.Description
   End If
Exit Function
End Function



Function InsertarDebbAtenciones(ByVal oTabla As DOAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   InsertarDebbAtenciones = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_conexion
       .CommandText = "debbAtencionesAgregar"
           Set oParameter = .CreateParameter("@IdTipoReferenciaDestino", adInteger, adParamInput, 0, IIf(oTabla.IdTipoReferenciaDestino = 0, Null, oTabla.IdTipoReferenciaDestino)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoReferenciaOrigen", adInteger, adParamInput, 0, IIf(oTabla.IdTipoReferenciaOrigen = 0, Null, oTabla.IdTipoReferenciaOrigen)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstablecimientoDestino", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimientoDestino = 0, Null, oTabla.IdEstablecimientoDestino)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstablecimientoOrigen", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimientoOrigen = 0, Null, oTabla.IdEstablecimientoOrigen)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraIngreso", adChar, adParamInput, 5, IIf(oTabla.HoraIngreso = "", Null, oTabla.HoraIngreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaIngreso", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaIngreso = 0, Null, oTabla.FechaIngreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoServicio", adInteger, adParamInput, 0, IIf(oTabla.IdTipoServicio = 0, Null, oTabla.IdTipoServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.idPaciente = 0, Null, oTabla.idPaciente)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, oTabla.idAtencion): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoCondicionALEstab", adInteger, adParamInput, 0, IIf(oTabla.IdTipoCondicionALEstab = 0, Null, oTabla.IdTipoCondicionALEstab)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@DireccionDomicilio", adVarChar, adParamInput, 50, IIf(oTabla.DireccionDomicilio = "", Null, oTabla.DireccionDomicilio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaEgresoAdministrativo", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaEgresoAdministrativo = 0, Null, oTabla.FechaEgresoAdministrativo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedicoRespNacimiento", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoRespNacimiento = 0, Null, oTabla.IdMedicoRespNacimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCamaEgreso", adInteger, adParamInput, 0, IIf(oTabla.IdCamaEgreso = 0, Null, oTabla.IdCamaEgreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCamaIngreso", adInteger, adParamInput, 0, IIf(oTabla.IdCamaIngreso = 0, Null, oTabla.IdCamaIngreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioEgreso", adInteger, adParamInput, 0, IIf(oTabla.IdServicioEgreso = 0, Null, oTabla.IdServicioEgreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoAlta", adInteger, adParamInput, 0, IIf(oTabla.IdTipoAlta = 0, Null, oTabla.IdTipoAlta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCondicionAlta", adInteger, adParamInput, 0, IIf(oTabla.IdCondicionAlta = 0, Null, oTabla.IdCondicionAlta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoEdad", adInteger, adParamInput, 0, IIf(oTabla.IdTipoEdad = 0, Null, oTabla.IdTipoEdad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@RecienNacido", adBoolean, adParamInput, 0, IIf(oTabla.RecienNacido = 0, Null, oTabla.RecienNacido)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NombreAcompaniante", adVarChar, adParamInput, 30, IIf(oTabla.NombreAcompaniante = "", Null, oTabla.NombreAcompaniante)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdOrigenAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdOrigenAtencion = 0, Null, oTabla.IdOrigenAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TieneNecropsia", adBoolean, adParamInput, 0, IIf(oTabla.TieneNecropsia = 0, Null, oTabla.TieneNecropsia)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDestinoAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdDestinoAtencion = 0, Null, oTabla.IdDestinoAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraEgresoAdministrativo", adChar, adParamInput, 5, IIf(oTabla.HoraEgresoAdministrativo = "", Null, oTabla.HoraEgresoAdministrativo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoCondicionAlServicio", adInteger, adParamInput, 0, IIf(oTabla.IdTipoCondicionAlServicio = 0, Null, oTabla.IdTipoCondicionAlServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Observacion", adVarChar, adParamInput, 200, IIf(oTabla.Observacion = "", Null, oTabla.Observacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraEgreso", adChar, adParamInput, 5, IIf(oTabla.HoraEgreso = "", Null, oTabla.HoraEgreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaEgreso", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaEgreso = 0, Null, oTabla.FechaEgreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedicoEgreso", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoEgreso = 0, Null, oTabla.IdMedicoEgreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstablecimientoNoMinsaDestino", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimientoNoMinsaDestino = 0, Null, oTabla.IdEstablecimientoNoMinsaDestino)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstablecimientoNoMinsaOrigen", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimientoNoMinsaOrigen = 0, Null, oTabla.IdEstablecimientoNoMinsaOrigen)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Edad", adInteger, adParamInput, 0, IIf(oTabla.Edad = 0, Null, oTabla.Edad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEspecialidadMedico", adInteger, adParamInput, 0, IIf(oTabla.IdEspecialidadMedico = 0, Null, oTabla.IdEspecialidadMedico)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedicoIngreso", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoIngreso = 0, Null, oTabla.IdMedicoIngreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioIngreso", adInteger, adParamInput, 0, IIf(oTabla.IdServicioIngreso = 0, Null, oTabla.IdServicioIngreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoGravedad", adInteger, adParamInput, 0, IIf(oTabla.IdTipoGravedad = 0, Null, oTabla.IdTipoGravedad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.idCuentaAtencion = 0, Null, oTabla.idCuentaAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HuboInfeccionIntraHospitalaria", adBoolean, adParamInput, 0, IIf(oTabla.HuboInfeccionIntraHospitalaria = 0, Null, oTabla.HuboInfeccionIntraHospitalaria)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idFormaPago", adInteger, adParamInput, 4, IIf(oTabla.IdFormaPago = 0, Null, oTabla.IdFormaPago)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idFuenteFinanciamiento", adInteger, adParamInput, 4, IIf(oTabla.IdFuenteFinanciamiento = 0, Null, oTabla.IdFuenteFinanciamiento)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idEstadoAtencion", adInteger, adParamInput, 4, oTabla.IdEstadoAtencion): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroReferenciaOrigen", adVarChar, adParamInput, 20, IIf(oTabla.NroReferenciaOrigen = "", Null, oTabla.NroReferenciaOrigen)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroReferenciaDestino", adVarChar, adParamInput, 20, IIf(oTabla.NroReferenciaDestino = "", Null, oTabla.NroReferenciaDestino)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EsPacienteExterno", adBoolean, adParamInput, 0, IIf(oTabla.EsPacienteExterno = True, 1, 0)): .Parameters.Append oParameter
       .Execute
   End With
   InsertarDebbAtenciones = True
Exit Function
ManejadorDeError:
      MsgBox Err.Number & " " + Err.Description
Exit Function
End Function

Function InsertarDebbAtencionDiagnostico(ByVal oTabla As DOAtencionDiagnostico) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   InsertarDebbAtencionDiagnostico = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_conexion
       .CommandText = "debbAtencionesDiagnosticosAgregar"
           Set oParameter = .CreateParameter("@IdSubclasificacionDx", adInteger, adParamInput, 0, IIf(oTabla.IdSubclasificacionDx = 0, Null, oTabla.IdSubclasificacionDx)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdClasificacionDx", adInteger, adParamInput, 0, IIf(oTabla.IdClasificacionDx = 0, Null, oTabla.IdClasificacionDx)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, Null, oTabla.IdDiagnostico)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencionDiagnostico", adInteger, adParamInput, 0, oTabla.IdAtencionDiagnostico): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@labConfHIS", adVarChar, adParamInput, 3, IIf(oTabla.labConfHIS = "", Null, oTabla.labConfHIS)): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
   End With
   InsertarDebbAtencionDiagnostico = True
Exit Function
ManejadorDeError:
      MsgBox Err.Number & " " + Err.Description
Exit Function
End Function


Function InsertarDebbAtencionEmergencia(ByVal oTabla As DOAtencionEmergencia) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   InsertarDebbAtencionEmergencia = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_conexion
       .CommandText = "debbAtencionesEmergenciaAgregar"
           Set oParameter = .CreateParameter("@IdTipoAgenteAGAN", adInteger, adParamInput, 0, IIf(oTabla.IdTipoAgenteAGAN = 0, Null, oTabla.IdTipoAgenteAGAN)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdGrupoOcupacionalALAB", adInteger, adParamInput, 0, IIf(oTabla.IdGrupoOcupacionalALAB = 0, Null, oTabla.IdGrupoOcupacionalALAB)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPosicionLesionadoALAB", adInteger, adParamInput, 0, IIf(oTabla.IdPosicionLesionadoALAB = 0, Null, oTabla.IdPosicionLesionadoALAB)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUbicacionLesionado", adInteger, adParamInput, 0, IIf(oTabla.IdUbicacionLesionado = 0, Null, oTabla.IdUbicacionLesionado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoTransporte", adInteger, adParamInput, 0, IIf(oTabla.IdTipoTransporte = 0, Null, oTabla.IdTipoTransporte)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoVehiculo", adInteger, adParamInput, 0, IIf(oTabla.IdTipoVehiculo = 0, Null, oTabla.IdTipoVehiculo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdClaseAccidente", adInteger, adParamInput, 0, IIf(oTabla.IdClaseAccidente = 0, Null, oTabla.IdClaseAccidente)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdRelacionAgresorVictima", adInteger, adParamInput, 0, IIf(oTabla.IdRelacionAgresorVictima = 0, Null, oTabla.IdRelacionAgresorVictima)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdSeguridad", adInteger, adParamInput, 0, IIf(oTabla.IdSeguridad = 0, Null, oTabla.IdSeguridad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoEvento", adInteger, adParamInput, 0, IIf(oTabla.IdTipoEvento = 0, Null, oTabla.IdTipoEvento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdLugarEvento", adInteger, adParamInput, 0, IIf(oTabla.IdLugarEvento = 0, Null, oTabla.IdLugarEvento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCausaExternaMorbilidad", adInteger, adParamInput, 0, IIf(oTabla.IdCausaExternaMorbilidad = 0, Null, oTabla.IdCausaExternaMorbilidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencionEmergencia", adInteger, adParamInput, 0, oTabla.IdAtencionEmergencia): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
   End With
   InsertarDebbAtencionEmergencia = True
Exit Function
ManejadorDeError:
      MsgBox Err.Number & " " + Err.Description
Exit Function
End Function


Function InsertarDebbEstanciaHospitalaria(ByVal oTabla As DOEstanciaHospitalaria) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   InsertarDebbEstanciaHospitalaria = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_conexion
       .CommandText = "debbAtencionesEstanciaHospitalariaAgregar"
           Set oParameter = .CreateParameter("@DiasEstancia", adDecimal, adParamInput, 5, IIf(oTabla.DiasEstancia = 0, Null, oTabla.DiasEstancia)):
           oParameter.Precision = 8
           oParameter.NumericScale = 2
           .Parameters.Append oParameter
           
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdFacturacionServicio", adInteger, adParamInput, 0, IIf(oTabla.IdFacturacionServicio = 0, Null, oTabla.IdFacturacionServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedicoOrdena", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoOrdena = 0, Null, oTabla.IdMedicoOrdena)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCama", adInteger, adParamInput, 0, IIf(oTabla.idCama = 0, Null, oTabla.idCama)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, IIf(oTabla.idServicio = 0, Null, oTabla.idServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraDesocupacion", adChar, adParamInput, 5, IIf(oTabla.HoraDesocupacion = "", Null, oTabla.HoraDesocupacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaDesocupacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaDesocupacion = 0, Null, oTabla.FechaDesocupacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraOcupacion", adChar, adParamInput, 5, IIf(oTabla.HoraOcupacion = "", Null, oTabla.HoraOcupacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaOcupacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaOcupacion = 0, Null, oTabla.FechaOcupacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Secuencia", adInteger, adParamInput, 0, IIf(oTabla.Secuencia = 0, Null, oTabla.Secuencia)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstanciaHospitalaria", adInteger, adParamInput, 0, oTabla.IdEstanciaHospitalaria): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@LlegoAlServicio", adInteger, adParamInput, 0, oTabla.LlegoAlServicio): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
   End With
   InsertarDebbEstanciaHospitalaria = True
Exit Function
ManejadorDeError:
      MsgBox Err.Number & " " + Err.Description
Exit Function
End Function


Function InsertarDebbAtencionNacimiento(ByVal oTabla As DOAtencionNacimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   InsertarDebbAtencionNacimiento = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_conexion
       .CommandText = "debbAtencionesNacimientosAgregar"
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCondicionRN", adInteger, adParamInput, 0, IIf(oTabla.idCondicionRN = 0, Null, oTabla.idCondicionRN)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoSexo", adInteger, adParamInput, 0, IIf(oTabla.idTipoSexo = 0, Null, oTabla.idTipoSexo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Peso", adDouble, adParamInput, 0, IIf(oTabla.Peso = 0, Null, oTabla.Peso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Talla", adDouble, adParamInput, 0, IIf(oTabla.Talla = 0, Null, oTabla.Talla)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EdadSemanas", adInteger, adParamInput, 0, IIf(oTabla.EdadSemanas = 0, Null, oTabla.EdadSemanas)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaNacimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaNacimiento = 0, Null, oTabla.FechaNacimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdNacimiento", adInteger, adParamInput, 0, oTabla.idNacimiento): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.idNacimiento = .Parameters("@IdNacimiento")
   End With
   InsertarDebbAtencionNacimiento = True
Exit Function
ManejadorDeError:
     MsgBox Err.Number & " " + Err.Description
Exit Function
End Function



Function InsertarDebbCita(ByVal oTabla As DOCita) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   InsertarDebbCita = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_conexion
       .CommandText = "debbCitasAgregar"
           Set oParameter = .CreateParameter("@HoraSolicitud", adChar, adParamInput, 5, IIf(oTabla.HoraSolicitud = "", Null, oTabla.HoraSolicitud)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaSolicitud", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaSolicitud = 0, Null, oTabla.FechaSolicitud)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProgramacion", adInteger, adParamInput, 0, IIf(oTabla.IdProgramacion = 0, Null, oTabla.IdProgramacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 4, IIf(oTabla.idServicio = 0, Null, oTabla.idServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraFin", adChar, adParamInput, 5, IIf(oTabla.HoraFin = "", Null, oTabla.HoraFin)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraInicio", adChar, adParamInput, 5, IIf(oTabla.HoraInicio = "", Null, oTabla.HoraInicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCita", adInteger, adParamInput, 0, oTabla.IdCita): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 4, IIf(oTabla.Fecha = 0, Null, oTabla.Fecha)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstadoCita", adInteger, adParamInput, 4, IIf(oTabla.IdEstadoCita = 0, Null, oTabla.IdEstadoCita)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 4, IIf(oTabla.idMedico = 0, Null, oTabla.idMedico)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, 4, IIf(oTabla.IdEspecialidad = 0, Null, oTabla.IdEspecialidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 4, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 4, IIf(oTabla.idPaciente = 0, Null, oTabla.idPaciente)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
   End With
   InsertarDebbCita = True
Exit Function
ManejadorDeError:
   MsgBox Err.Number & " " + Err.Description
Exit Function
End Function


Function InsertarDebbMovimientoHistoriaClinica(ByVal oTabla As DOMovimientoHistoriaClinica) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   InsertarDebbMovimientoHistoriaClinica = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_conexion
       .CommandText = "debbMovimientosHistoriaClinicaAgregar"
           Set oParameter = .CreateParameter("@NroFolios", adInteger, adParamInput, 0, IIf(oTabla.NroFolios = 0, Null, oTabla.NroFolios)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioDestino", adInteger, adParamInput, 0, IIf(oTabla.idServicioDestino = 0, Null, oTabla.idServicioDestino)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioOrigen", adInteger, adParamInput, 0, IIf(oTabla.IdServicioOrigen = 0, Null, oTabla.IdServicioOrigen)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Observacion", adVarChar, adParamInput, 100, IIf(oTabla.Observacion = "", Null, oTabla.Observacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMotivo", adInteger, adParamInput, 0, IIf(oTabla.IdMotivo = 0, Null, oTabla.IdMotivo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaMovimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaMovimiento = 0, Null, oTabla.FechaMovimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.idPaciente = 0, Null, oTabla.idPaciente)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, oTabla.IdMovimiento): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEmpleadoRecepcion", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleadoRecepcion = 0, Null, oTabla.IdEmpleadoRecepcion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEmpleadoTransporte", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleadoTransporte = 0, Null, oTabla.IdEmpleadoTransporte)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEmpleadoArchivo", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleadoArchivo = 0, Null, oTabla.IdEmpleadoArchivo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdGrupoMovimiento", adInteger, adParamInput, 0, IIf(oTabla.IdGrupoMovimiento = 0, Null, oTabla.IdGrupoMovimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter

       .Execute
   End With
   InsertarDebbMovimientoHistoriaClinica = True
Exit Function
ManejadorDeError:
   MsgBox Err.Number & " " + Err.Description
Exit Function
End Function

Function InsertarDebbHistorias(ByVal oTabla As DOHistoriaClinica) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   InsertarDebbHistorias = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_conexion
       .CommandText = "HistoriasClinicasAgregar"
           Set oParameter = .CreateParameter("@IdTipoNumeracionAnterior", adInteger, adParamInput, 0, IIf(oTabla.IdTipoNumeracionAnterior = 0, Null, oTabla.IdTipoNumeracionAnterior)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroHistoriaClinicaAnterior", adInteger, adParamInput, 0, IIf(oTabla.NroHistoriaClinicaAnterior = 0, Null, oTabla.NroHistoriaClinicaAnterior)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoNumeracion", adInteger, adParamInput, 0, IIf(oTabla.IdTipoNumeracion = 0, Null, oTabla.IdTipoNumeracion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, IIf(oTabla.NroHistoriaClinica = 0, Null, oTabla.NroHistoriaClinica)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaPasoAPasivo", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaPasoAPasivo = 0, Null, oTabla.FechaPasoAPasivo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoHistoria", adInteger, adParamInput, 4, IIf(oTabla.IdTipoHistoria = 0, Null, oTabla.IdTipoHistoria)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstadoHistoria", adInteger, adParamInput, 4, IIf(oTabla.IdEstadoHistoria = 0, Null, oTabla.IdEstadoHistoria)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 4, IIf(oTabla.idPaciente = 0, Null, oTabla.idPaciente)): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
          ' oTabla.NroHistoriaClinica = .Parameters("@NroHistoriaClinica")
   End With
 
   InsertarDebbHistorias = True
 
Exit Function
ManejadorDeError:
   ml_Errores = Err.Number & " " & Err.Description
Exit Function
End Function

