VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FarmVentas 
   Caption         =   "Form1"
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13830
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FarmVentas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9420
   ScaleWidth      =   13830
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CargaInventarioExcel 
      Enabled         =   0   'False
      Height          =   315
      Left            =   13290
      Picture         =   "FarmVentas.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   7575
      Width           =   435
   End
   Begin VB.TextBox txtHtotCantidad 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   10350
      TabIndex        =   61
      Top             =   7605
      Visible         =   0   'False
      Width           =   1125
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3000
      Left            =   15
      TabIndex        =   7
      Top             =   30
      Width           =   13800
      _ExtentX        =   24342
      _ExtentY        =   5292
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
      TabCaption(0)   =   "Datos Basicos de Cabecera"
      TabPicture(0)   =   "FarmVentas.frx":110C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label13"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblNcuenta"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label14"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label10"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label8"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label19"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label21"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblOrdenPago"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtFprescribe"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtFregistro"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtNpreventa"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtDatosDeCuenta"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtNcuenta"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtTipoComprobante"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cmbTipoReceta"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmbTipoFinanciamiento"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cmbAlmOrigen"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtDocumento"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtHoraRegistro"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtEstado"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "chkPlanNoCubre"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "fraTipoVenta"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtNreceta"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "cmdPaquetes"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "cmbPrescriptor"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtNombrePaciente"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtNhistoria"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtPlan"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "chkHistorico"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "cmdBuscaCuentaPorApellidos"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "btnBuscarPaciente"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "cmbBuscaReceta"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "cmdStockMinimo"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).ControlCount=   39
      TabCaption(1)   =   "Datos complementarios de Cabecera"
      TabPicture(1)   =   "FarmVentas.frx":1128
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label15"
      Tab(1).Control(1)=   "Label12"
      Tab(1).Control(2)=   "Label11"
      Tab(1).Control(3)=   "Label20"
      Tab(1).Control(4)=   "Label9"
      Tab(1).Control(5)=   "Label7"
      Tab(1).Control(6)=   "txtCaja"
      Tab(1).Control(7)=   "txtCajero"
      Tab(1).Control(8)=   "txtVendedor"
      Tab(1).Control(9)=   "txtTurno"
      Tab(1).Control(10)=   "txtDx"
      Tab(1).Control(11)=   "txtNombreDx"
      Tab(1).Control(12)=   "txtObservaciones"
      Tab(1).Control(13)=   "cmdBuscaDx"
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "Despacho para FARMACIA UNIDOSIS"
      TabPicture(2)   =   "FarmVentas.frx":1144
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FraUnidosis"
      Tab(2).ControlCount=   1
      Begin VB.CommandButton cmdStockMinimo 
         Height          =   300
         Left            =   5295
         Picture         =   "FarmVentas.frx":1160
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Lista de ITEMS que están debajo de su STOCK MINIMO"
         Top             =   780
         Width           =   300
      End
      Begin VB.Frame FraUnidosis 
         Height          =   2265
         Left            =   -74850
         TabIndex        =   67
         Top             =   480
         Width           =   13470
         Begin VB.ComboBox cmbUnidosis 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2100
            TabIndex        =   70
            Top             =   960
            Visible         =   0   'False
            Width           =   4215
         End
         Begin VB.TextBox txtCodigoDespacho 
            Height          =   315
            Left            =   2100
            MaxLength       =   30
            TabIndex        =   69
            Top             =   615
            Width           =   1245
         End
         Begin VB.ComboBox cmbFarmaciaOrigen 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2100
            TabIndex        =   68
            Top             =   240
            Width           =   4215
         End
         Begin VB.Label lblUnidosis 
            AutoSize        =   -1  'True
            Caption         =   "F.UNIDOSIS (destino)"
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   165
            TabIndex        =   73
            Top             =   1005
            Visible         =   0   'False
            Width           =   1785
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Código de despacho"
            Height          =   210
            Left            =   165
            TabIndex        =   72
            Top             =   645
            Width           =   1665
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Farmacia (origen)"
            Height          =   210
            Left            =   165
            TabIndex        =   71
            Top             =   240
            Width           =   1410
         End
      End
      Begin VB.CommandButton cmbBuscaReceta 
         Height          =   330
         Left            =   13410
         Picture         =   "FarmVentas.frx":16EA
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   1710
         Width           =   300
      End
      Begin VB.CommandButton cmdBuscaDx 
         Height          =   330
         Left            =   -72450
         Picture         =   "FarmVentas.frx":1C74
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   1695
         Width           =   360
      End
      Begin VB.CommandButton btnBuscarPaciente 
         Height          =   330
         Left            =   2310
         Picture         =   "FarmVentas.frx":21FE
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   2490
         Width           =   300
      End
      Begin VB.CommandButton cmdBuscaCuentaPorApellidos 
         Height          =   330
         Left            =   2325
         Picture         =   "FarmVentas.frx":2788
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   1650
         Width           =   300
      End
      Begin VB.CheckBox chkHistorico 
         Alignment       =   1  'Right Justify
         Caption         =   "Consumo histórico"
         Height          =   225
         Left            =   8745
         TabIndex        =   59
         Top             =   2550
         Width           =   1785
      End
      Begin VB.TextBox txtPlan 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7905
         TabIndex        =   57
         Top             =   2070
         Width           =   5820
      End
      Begin VB.TextBox txtNhistoria 
         Alignment       =   1  'Right Justify
         Height          =   324
         Left            =   1065
         MaxLength       =   9
         TabIndex        =   56
         ToolTipText     =   "Ingrese el Nro de Historia Clínica"
         Top             =   2490
         Width           =   1245
      End
      Begin VB.TextBox txtNombrePaciente 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2655
         MaxLength       =   100
         TabIndex        =   55
         Top             =   2490
         Width           =   3450
      End
      Begin VB.ComboBox cmbPrescriptor 
         Height          =   312
         Left            =   1065
         TabIndex        =   51
         Top             =   2070
         Width           =   3735
      End
      Begin VB.CommandButton cmdPaquetes 
         Caption         =   "Carga Paquetes"
         Height          =   435
         Left            =   11940
         TabIndex        =   49
         Top             =   750
         Width           =   1755
      End
      Begin VB.TextBox txtNreceta 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   12495
         MaxLength       =   30
         TabIndex        =   48
         Top             =   1695
         Width           =   915
      End
      Begin VB.Frame fraTipoVenta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1080
         TabIndex        =   43
         Top             =   1080
         Width           =   8580
         Begin Threed.SSOption optVentas 
            Height          =   285
            Left            =   75
            TabIndex        =   44
            Top             =   180
            Width           =   2820
            _ExtentX        =   4974
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
            Caption         =   "VENTA (tienen algún Seguro)"
            Value           =   -1
         End
         Begin Threed.SSOption optPreventa 
            Height          =   285
            Left            =   6270
            TabIndex        =   45
            Top             =   180
            Width           =   2250
            _ExtentX        =   3969
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
            Caption         =   "PRE-VENTA (Pagantes)"
         End
         Begin Threed.SSOption optVtaSinPlan 
            Height          =   285
            Left            =   3000
            TabIndex        =   46
            Top             =   180
            Visible         =   0   'False
            Width           =   2925
            _ExtentX        =   5159
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
            Caption         =   "VENTA (sin Fte.Financiamiento)"
         End
      End
      Begin VB.CheckBox chkPlanNoCubre 
         Alignment       =   1  'Right Justify
         Caption         =   "IAFA NO cubre"
         Height          =   345
         Left            =   6180
         TabIndex        =   42
         Top             =   2490
         Width           =   1590
      End
      Begin VB.TextBox txtEstado 
         Enabled         =   0   'False
         Height          =   324
         Left            =   8685
         MaxLength       =   30
         TabIndex        =   40
         Top             =   375
         Width           =   1785
      End
      Begin VB.TextBox txtHoraRegistro 
         Enabled         =   0   'False
         Height          =   324
         Left            =   12915
         MaxLength       =   30
         TabIndex        =   37
         Top             =   390
         Width           =   795
      End
      Begin VB.TextBox txtDocumento 
         Enabled         =   0   'False
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
         Left            =   1065
         MaxLength       =   30
         TabIndex        =   28
         Top             =   420
         Width           =   1665
      End
      Begin VB.ComboBox cmbAlmOrigen 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   1080
         TabIndex        =   27
         Top             =   765
         Width           =   4215
      End
      Begin VB.ComboBox cmbTipoFinanciamiento 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9060
         TabIndex        =   26
         Top             =   1680
         Width           =   2580
      End
      Begin VB.ComboBox cmbTipoReceta 
         Height          =   312
         Left            =   11505
         TabIndex        =   25
         Top             =   1245
         Width           =   2220
      End
      Begin VB.TextBox txtTipoComprobante 
         Enabled         =   0   'False
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
         Left            =   2730
         TabIndex        =   24
         Top             =   420
         Width           =   2505
      End
      Begin VB.TextBox txtNcuenta 
         Height          =   324
         Left            =   1065
         MaxLength       =   30
         TabIndex        =   23
         Top             =   1665
         Width           =   1245
      End
      Begin VB.TextBox txtDatosDeCuenta 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2655
         TabIndex        =   22
         Top             =   1665
         Width           =   5160
      End
      Begin VB.TextBox txtNpreventa 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6660
         TabIndex        =   21
         Top             =   390
         Width           =   1215
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   324
         Left            =   -73620
         MaxLength       =   100
         TabIndex        =   14
         Top             =   2055
         Width           =   4995
      End
      Begin VB.TextBox txtNombreDx 
         Enabled         =   0   'False
         Height          =   324
         Left            =   -72090
         TabIndex        =   13
         Top             =   1695
         Width           =   3465
      End
      Begin VB.TextBox txtDx 
         Height          =   324
         Left            =   -73605
         MaxLength       =   30
         TabIndex        =   12
         ToolTipText     =   "Ingrese el Dx (4 digitos)"
         Top             =   1695
         Width           =   1095
      End
      Begin VB.TextBox txtTurno 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   -73605
         TabIndex        =   11
         Top             =   690
         Width           =   1785
      End
      Begin VB.TextBox txtVendedor 
         Height          =   324
         Left            =   -73605
         TabIndex        =   10
         Top             =   1350
         Width           =   4995
      End
      Begin VB.TextBox txtCajero 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   -73605
         TabIndex        =   9
         Top             =   1020
         Width           =   4995
      End
      Begin VB.TextBox txtCaja 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   -70455
         TabIndex        =   8
         Top             =   660
         Width           =   1845
      End
      Begin MSMask.MaskEdBox txtFregistro 
         Height          =   315
         Left            =   11535
         TabIndex        =   38
         Top             =   390
         Width           =   1350
         _ExtentX        =   2381
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
      Begin MSMask.MaskEdBox txtFprescribe 
         Height          =   315
         Left            =   6000
         TabIndex        =   53
         Top             =   2085
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label lblOrdenPago 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "N° Orden de Pago"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   12150
         TabIndex        =   66
         Top             =   2580
         Width           =   1515
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "F.Prescripción"
         Height          =   210
         Left            =   4905
         TabIndex        =   54
         Top             =   2145
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Prescriptor"
         Height          =   225
         Left            =   120
         TabIndex        =   52
         Top             =   2145
         Width           =   870
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "N°Receta"
         Height          =   210
         Left            =   11700
         TabIndex        =   50
         Top             =   1725
         Width           =   780
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   210
         Left            =   8100
         TabIndex        =   41
         Top             =   405
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "F.Registro"
         Height          =   210
         Left            =   10695
         TabIndex        =   39
         Top             =   420
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Documento"
         Height          =   210
         Left            =   120
         TabIndex        =   36
         Top             =   450
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Farmacia"
         Height          =   210
         Left            =   120
         TabIndex        =   35
         Top             =   825
         Width           =   690
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Receta"
         Height          =   210
         Left            =   10425
         TabIndex        =   34
         Top             =   1275
         Width           =   990
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Paciente"
         Height          =   210
         Left            =   120
         TabIndex        =   33
         Top             =   2490
         Width           =   705
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Producto/Plan"
         Height          =   210
         Left            =   7905
         TabIndex        =   32
         Top             =   1725
         Width           =   1155
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Venta"
         Height          =   210
         Left            =   120
         TabIndex        =   31
         Top             =   1260
         Width           =   915
      End
      Begin VB.Label lblNcuenta 
         AutoSize        =   -1  'True
         Caption         =   "N° Cuenta"
         Height          =   210
         Left            =   120
         TabIndex        =   30
         Top             =   1695
         Width           =   855
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "N° PreVenta"
         Height          =   210
         Left            =   5550
         TabIndex        =   29
         Top             =   420
         Width           =   1035
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Diagnóstico"
         Height          =   210
         Left            =   -74805
         TabIndex        =   20
         Top             =   1740
         Width           =   930
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor"
         Height          =   210
         Left            =   -74790
         TabIndex        =   19
         Top             =   1386
         Width           =   810
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         Height          =   210
         Left            =   -74805
         TabIndex        =   18
         Top             =   2085
         Width           =   1170
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Turno"
         Height          =   210
         Left            =   -74790
         TabIndex        =   17
         Top             =   750
         Width           =   495
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Cajero"
         Height          =   210
         Left            =   -74790
         TabIndex        =   16
         Top             =   1050
         Width           =   510
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Caja"
         Height          =   210
         Left            =   -70845
         TabIndex        =   15
         Top             =   690
         Width           =   330
      End
   End
   Begin SighFarmacia.ucVentas grdProductos 
      Height          =   3465
      Left            =   0
      TabIndex        =   6
      Top             =   3060
      Width           =   13785
      _ExtentX        =   24315
      _ExtentY        =   6112
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   0
      TabIndex        =   1
      Top             =   6600
      Width           =   13815
      Begin VB.CommandButton btnImprimeFichaSIS 
         Caption         =   "FUA"
         Height          =   700
         Left            =   12375
         Picture         =   "FarmVentas.frx":2D12
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   135
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.CommandButton btnImprimir 
         Caption         =   "Imprime"
         Height          =   700
         Left            =   1515
         Picture         =   "FarmVentas.frx":31EB
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   165
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.CommandButton btnNuevo 
         Caption         =   "Nueva Venta (F3)"
         DisabledPicture =   "FarmVentas.frx":36C4
         DownPicture     =   "FarmVentas.frx":3B88
         Height          =   700
         Left            =   90
         Picture         =   "FarmVentas.frx":4074
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   150
         Width           =   1365
      End
      Begin VB.Frame FraRedondeo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   8970
         TabIndex        =   3
         Top             =   210
         Visible         =   0   'False
         Width           =   2925
         Begin VB.TextBox txtRedondeo 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1590
            MaxLength       =   30
            TabIndex        =   4
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Redondear Total"
            Height          =   210
            Left            =   120
            TabIndex        =   5
            Top             =   180
            Width           =   1365
         End
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "FarmVentas.frx":4560
         DownPicture     =   "FarmVentas.frx":4A24
         Height          =   700
         Left            =   7002
         Picture         =   "FarmVentas.frx":4F10
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   135
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "FarmVentas.frx":53FC
         DownPicture     =   "FarmVentas.frx":585C
         Height          =   700
         Left            =   5464
         Picture         =   "FarmVentas.frx":5CD1
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   135
         Width           =   1365
      End
   End
   Begin UltraGrid.SSUltraGrid grdHistorico 
      Height          =   1815
      Left            =   0
      TabIndex        =   58
      Top             =   7590
      Visible         =   0   'False
      Width           =   13755
      _ExtentX        =   24262
      _ExtentY        =   3201
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   71303188
      BorderStyle     =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   "FarmVentas.frx":6146
      Caption         =   "Consumo histórico del PACIENTE x CUENTA"
   End
End
Attribute VB_Name = "FarmVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenmiento de ventas a Pacientes
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ReporteUtil As New ReporteUtil
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim mi_Opcion As sghOpciones
Dim ml_idUsuario As Long
Dim ml_movNumero As String
Dim ml_IdTipoVentaSeleccionada As Long          '0=VentaDirecta      1=PreVenta
Dim mo_cmbFarmaciaOrigen As New SIGHEntidades.ListaDespleglable
Dim mo_cmbAlmacenOrigen As New SIGHEntidades.ListaDespleglable
Dim mo_cmbUnidosis As New SIGHEntidades.ListaDespleglable
Dim mo_cmbPrescriptor As New SIGHEntidades.ListaDespleglable
Dim mo_cmbTipoFinanciamiento As New SIGHEntidades.ListaDespleglable
Dim mo_cmbTipoReceta As New SIGHEntidades.ListaDespleglable
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminAdmision As New ReglasAdmision
Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasSISgalenhos As New SIGHSis.ReglasSISgalenhos
Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
Dim mo_sighProxies As New SIGHProxies.Procesos
Dim oRsTipoFinanciamiento As New Recordset
Dim oRsItemsUnidosis As New Recordset
Dim ms_MensajeError As String
Dim ml_IdTipoComprobante As Long
Dim ml_IdDiagnostico As Long
Dim ml_IdPaciente As Long
Dim ml_IdVendedor As Long
Dim ml_IdCajero As Long
Dim mo_DofarmPreVenta As New DoFarmPreVenta
Dim mo_DoPaciente As New DOPaciente
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lnTotalDocumento As Double
Dim mRs_Productos As New Recordset
Dim mo_farmMovimiento1 As New DoFarmMovimiento
Dim mo_DoFarmMovimiento As New sighComun.DoFarmMovimiento
Dim mo_DoFarmMovimientoVentas As New sighComun.DoFarmMovimientoVentas
Dim ml_idTipoConcepto As Long
Const lcConstanteMovimientoSalida As String = "S"
Const lcConstantePreVenta As String = "P"
Const lcConstanteVentaDirecta As String = "D"
Const lcConstanteMovimientoEntrada As String = "E"
Dim ml_ElTipoFinanciamientoSeUsaEnFarmacia As Boolean
Dim lcPosicionDefaultCombo As String
Dim ml_idFuenteFinanciamiento As Long
Dim ml_IdFuenteFinanciamientoDespacho As Long
Dim lnIdTipoServicio As Long
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Const lcEFE As String = "F"   'debb-16/02/2011
Dim lnIdReceta As Long
Dim wxParametro302 As String
Dim wxParametro509 As String
Dim wxParametro208 As String
Dim wxParametro280 As String
Dim wxParametro578 As String
Dim lnEpsPorcentaje As Double
Dim lbLaFarmaciaEsUnidosis As Boolean, lbLaFuenteFinanciamientoUsadoEnFUnidosis As Boolean, lnCuentaUnidosis As Long
Dim lbCuentaDeEmergenciaCerrada As Boolean

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
Property Let movNumero(lValue As String)
   ml_movNumero = lValue
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Let Opcion(lValue As sghOpciones)
   mi_Opcion = lValue
End Property
Property Let TipoVentaSeleccionada(lValue As Long)
   ml_IdTipoVentaSeleccionada = lValue
End Property



Function MuestraMensajeDeMantenimientoOK() As String
   Select Case mi_Opcion
   Case sghAgregar
        If optPreventa.Value Then
           MuestraMensajeDeMantenimientoOK = "Se agregó correctamente  la Preventa" + Chr(13) + Chr(13) + "Tiene que ir a CAJA a cancelar con el N° " + txtNpreventa.Text
        Else
           MuestraMensajeDeMantenimientoOK = "Se agregó correctamente  el Documento  N° " + txtDocumento.Text
        End If
   Case sghModificar
        If optPreventa.Value Then
           MuestraMensajeDeMantenimientoOK = "Se Modificó correctamente  la Preventa" + Chr(13) + Chr(13) + "Tiene que ir a CAJA a cancelar con el N° " + txtNpreventa.Text
        Else
           MuestraMensajeDeMantenimientoOK = "Se Modificó correctamente  el Documento  N° " + txtDocumento.Text
        End If
   Case sghEliminar
        If optPreventa.Value Then
           MuestraMensajeDeMantenimientoOK = "Se Anuló correctamente  la Preventa  N° " + txtNpreventa.Text
        Else
           MuestraMensajeDeMantenimientoOK = "Se Anuló correctamente  el Documento  N° " + txtDocumento.Text
        End If
   End Select
End Function

Private Sub btnAceptar_Click()
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   If mo_ReglasFarmacia.LaFarmaciaEstaRegenerandoSaldos(Val(mo_cmbAlmacenOrigen.BoundText)) = True Then
      btnCancelar_Click
      Exit Sub
   End If
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
            CargaDatosAlObjetosDeDatos
            If AgregarDatos() Then
                'debb-09/07/2015 (inicio)
                If optPreventa.Value Then
                   MsgBox MuestraMensajeDeMantenimientoOK, vbInformation, Me.Caption
                ElseIf Trim(lcBuscaParametro.SeleccionaFilaParametro(361)) = "S" Then
                   ImprimeDocumento
                End If
                If Me.optVentas And lblOrdenPago.Caption <> "" Then
                   MsgBox lblOrdenPago.Caption, vbInformation, ""
                End If
                'debb-09/07/2015  (fin)
                btnAceptar.Enabled = False
            Else
                MsgBox "No se pudo agregar los datos " + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
                grdProductos.RefrescaSaldos
            End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
            CargaDatosAlObjetosDeDatos
            If ModificarDatos() Then
                'debb-09/07/2015  (inicio)
                If optPreventa.Value Then
                   MsgBox MuestraMensajeDeMantenimientoOK, vbInformation, Me.Caption
                ElseIf Trim(lcBuscaParametro.SeleccionaFilaParametro(361)) = "S" Then
                   ImprimeDocumento
                End If
                'debb-09/07/2015   (fin)
                If Me.optVentas And lblOrdenPago.Caption <> "" Then
                   MsgBox lblOrdenPago.Caption, vbInformation, ""
                End If
                Me.Visible = False
                LimpiarVariablesDeMemoria
                LimpiarVariablesDeMemoria
            Else
                MsgBox "No se pudo modificar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
                grdProductos.RefrescaSaldos
            End If
       End If
   Case sghEliminar
        If MsgBox("Esta seguro de Anular ?", vbQuestion + vbYesNo, "") = vbYes Then
            CargaDatosAlObjetosDeDatos
            If Anular() Then
                MsgBox MuestraMensajeDeMantenimientoOK, vbInformation, Me.Caption
                Me.Visible = False
                LimpiarVariablesDeMemoria
            Else
                MsgBox "No se pudo eliminar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
            End If
        End If
   End Select
End Sub






Private Sub btnImprimeFichaSIS_Click()
    If mi_Opcion = sghAgregar Then
       Exit Sub
    End If
    Dim lnIdServicioActual As Long
    lnIdServicioActual = mo_DoFarmMovimientoVentas.IdServicioPaciente
    Dim oRsTmp As New Recordset
    Set oRsTmp = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(Val(txtNcuenta.Text))
    If oRsTmp.RecordCount > 0 Then
       If oRsTmp!IdTipoServicio <> 1 And Not IsNull(oRsTmp!idServicioEgreso) Then
          lnIdServicioActual = oRsTmp!idServicioEgreso
       End If
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
    
    
    Dim ml_FuaTipoAnexo2015 As Integer
    Dim oFua As New SIGHSis.clFUA
    oFua.SoloImprimeFUAyaGrabado = True
    oFua.lcNombrePc = mo_lcNombrePc
    oFua.lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroCitaCE
    oFua.idUsuario = SIGHEntidades.Usuario
    oFua.Opcion = sghConsultar
    oFua.idCuentaAtencion = Val(txtNcuenta.Text)
    oFua.IdServicio = lnIdServicioActual
    oFua.MostrarFormulario
    Set oFua = Nothing

End Sub

Private Sub btnImprimir_Click()
    ImprimeDocumento
End Sub

Private Sub btnNuevo_Click()
    If btnNuevo.Visible = True Then
        btnAceptar.Enabled = True
        LimpiarDatos
    End If
End Sub

Private Sub CargaInventarioExcel_Click()
    Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
    mo_ReglasReportes.ExportarRecordSetAexcelFast Me.grdHistorico.DataSource, Me.grdHistorico.Caption, "", "", Me.hwnd, , True
    Set mo_ReglasReportes = Nothing

End Sub

'kike 2017
Private Sub chkHistorico_Click()
    If chkHistorico.Value = 1 Then
        If Me.txtNcuenta.Text <> "" Then
             MuestraHistoricoDelPacienteXcuenta
        End If
    Else
        Me.Height = 8085
    End If
End Sub

'kike 2017
Sub MuestraHistoricoDelPacienteXcuenta()
    If chkHistorico.Value = 1 Then
            grdHistorico.Caption = "Consumo histórico del PACIENTE x CUENTA"
            Dim oRepConsumoPorCuenta As New RepConsumoPorCuenta, lnTotal1 As Double
            Dim oRsHistoricos As New Recordset
            Set oRsHistoricos = oRepConsumoPorCuenta.ProcesaConsumoXcuenta(Val(txtNcuenta.Text), lnTotal1, True, False)
            If mi_Opcion <> sghAgregar Then
               oRsHistoricos.Filter = "movNumero<>'" & Me.txtDocumento.Text & "'"
            End If
            Set grdHistorico.DataSource = oRsHistoricos
            grdHistorico.Visible = True
            txtHtotCantidad.Visible = True
            Me.Height = 9930
            Set oRepConsumoPorCuenta = Nothing
            Set oRsHistoricos = Nothing
    End If
End Sub

Private Sub chkPlanNoCubre_Click()
    If chkPlanNoCubre.Value = 1 Then
       mo_cmbTipoFinanciamiento.BoundText = 1   'contado
       grdProductos.LimpiarGrilla
       grdProductos.AgregaRegistro
    Else
       txtNcuenta_LostFocus
       grdProductos.LimpiarGrilla
    End If

End Sub

Private Sub cmbAlmOrigen_Click()
    grdProductos.IdAlmacen = Val(mo_cmbAlmacenOrigen.BoundText)
    'BlanquedaVariablesUnidosis
    If mi_Opcion = sghAgregar Then
       LimpiarDatos
    End If
    lbLaFarmaciaEsUnidosis = mo_ReglasFarmacia.FarmaciaEsUnidosis(Val(mo_cmbAlmacenOrigen.BoundText))
End Sub

Private Sub cmbAlmOrigen_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbAlmOrigen

End Sub



Private Sub cmbBuscaReceta_Click()
    Dim oBusqueda As New SIGHNegocios.clBuscaReceta
    oBusqueda.IdPuntoCarga = sghPtoCargaFarmacia
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
       txtNreceta.Text = oBusqueda.IdRecetaSeleccionada
       txtNreceta_LostFocus
    End If
    Set oBusqueda = Nothing
End Sub




Private Sub cmbFarmaciaOrigen_Click()
    mo_cmbAlmacenOrigen.BoundText = mo_cmbFarmaciaOrigen.BoundText
End Sub

Private Sub cmbFarmaciaOrigen_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, cmbFarmaciaOrigen
End Sub

Private Sub cmbPrescriptor_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbPrescriptor
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbTipoFinanciamiento_Click()
    If mo_cmbTipoFinanciamiento.BoundText = "" Then
       Exit Sub
    End If
    '
    btnBuscarPaciente.Enabled = False
    If mo_cmbTipoFinanciamiento.BoundText = "6" Or mo_cmbTipoFinanciamiento.BoundText = "10" Then  'defensa nacional, credito personal
       btnBuscarPaciente.Enabled = True
    End If
    '
    If oRsTipoFinanciamiento.RecordCount > 0 Then
        oRsTipoFinanciamiento.MoveFirst
        oRsTipoFinanciamiento.Find "idTipoFinanciamiento=" & mo_cmbTipoFinanciamiento.BoundText
        If Not oRsTipoFinanciamiento.EOF Then
           txtTipoComprobante.Text = oRsTipoFinanciamiento.Fields!dComprobante
           ml_IdTipoComprobante = oRsTipoFinanciamiento.Fields!idCajaTiposComprobante
           ml_ElTipoFinanciamientoSeUsaEnFarmacia = IIf(oRsTipoFinanciamiento.Fields!esFarmacia = True, True, False)
           grdProductos.IdTipoFinanciamiento = oRsTipoFinanciamiento.Fields!IdTipoFinanciamiento
           grdProductos.ElTipoFinanciamientoSeUsaEnFarmacia = ml_ElTipoFinanciamientoSeUsaEnFarmacia
           If Val(Me.txtNcuenta.Text) = 0 Then
               ml_idTipoConcepto = oRsTipoFinanciamiento.Fields!idTipoConcepto
           End If
        Else
           txtTipoComprobante.Text = ""
           ml_IdTipoComprobante = 0
        End If
    End If
    '
    On Error Resume Next
    cmbTipoReceta.SetFocus
End Sub



Private Sub cmbTipoFinanciamiento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbTipoFinanciamiento
  
End Sub



Private Sub cmbTipoReceta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbTipoReceta
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbTipoReceta_LostFocus()
    grdProductos.TabEnDescripcion
End Sub



Private Sub cmbUnidosis_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, cmbUnidosis
End Sub

Private Sub cmdBuscaCuentaPorApellidos_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaPacientes
    Dim oDOPaciente As New DOPaciente
    Dim oConexion As New Connection
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oBusqueda.TipoFiltro = sghFiltrarTodos
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(oBusqueda.IdRegistroSeleccionado, oConexion)
        If Not oDOPaciente Is Nothing Then
            ml_IdPaciente = oDOPaciente.IdPaciente
            txtNhistoria.Text = oDOPaciente.NroHistoriaClinica
            txtNombrePaciente.Text = Trim(oDOPaciente.ApellidoPaterno) + " " + Trim(oDOPaciente.ApellidoMaterno) + " " + oDOPaciente.PrimerNombre
            Dim oRsTmp As New Recordset
            Set oRsTmp = mo_ReglasFarmacia.FacturacionCuentasAtencionSeleccionarPorIdPaciente(ml_IdPaciente, oConexion, True)
            If oRsTmp.RecordCount > 0 Then
               txtNcuenta.Text = oRsTmp.Fields!idCuentaAtencion
            End If
            oRsTmp.Close
            Set oRsTmp = Nothing
            txtNcuenta_LostFocus
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
End Sub



Private Sub cmdPaquetes_Click()
    If cmbAlmOrigen.Text = "" Then
       MsgBox "Debe elegir Almacén", vbInformation, Me.Caption
       Exit Sub
    End If
    If optVentas.Value = True And txtDatosDeCuenta.Text = "" Then
       MsgBox "Debe ingresar un N° Cuenta", vbInformation, Me.Caption
       Exit Sub
    End If
    Dim oPaquetesBuscar As New SIGHNegocios.BuscaPaquetes
    Dim lnIdFactPaquete As Long
    oPaquetesBuscar.DebeConsiderarPaquete = sghTipoPaqueteSoloFarmacia
    oPaquetesBuscar.MostrarFormulario
    If oPaquetesBuscar.BotonPresionado = sghAceptar Then
       lnIdFactPaquete = oPaquetesBuscar.IdFactPaquete
       grdProductos.cargaPaqueteElegido lnIdFactPaquete
    End If
    Set oPaquetesBuscar = Nothing
End Sub

Private Sub Combo1_Change()

End Sub


Private Sub cmdStockMinimo_Click()
    CargaItemsDebajoDeStockMinimo
End Sub

Private Sub Form_Activate()
   If mi_Opcion <> sghAgregar Then
      On Error Resume Next
      cmbTipoReceta.SetFocus
   End If
   If mo_ReglasFarmacia.LaFarmaciaEstaRegenerandoSaldos(Val(mo_cmbAlmacenOrigen.BoundText)) = True Then
        btnCancelar_Click
        Exit Sub
   End If
End Sub

Private Sub Form_Initialize()
    Set mo_cmbAlmacenOrigen.MiComboBox = cmbAlmOrigen
    Set mo_cmbPrescriptor.MiComboBox = cmbPrescriptor
    Set mo_cmbTipoFinanciamiento.MiComboBox = cmbTipoFinanciamiento
    Set mo_cmbTipoReceta.MiComboBox = cmbTipoReceta
    Set mo_cmbUnidosis.MiComboBox = cmbUnidosis
End Sub

Private Sub Form_Load()
    SIGHEntidades.ParaAuditoriaPorCadaDato sghAudLimpiar, ""
    
    
    
    SSTab1.Tab = 0
    lblOrdenPago.Caption = ""
    txtFprescribe.Text = lcBuscaParametro.RetornaFechaHoraServidorSQL
    wxParametro208 = lcBuscaParametro.SeleccionaFilaParametro(208)
    wxParametro280 = lcBuscaParametro.SeleccionaFilaParametro(280)
    wxParametro578 = lcBuscaParametro.SeleccionaFilaParametro(578)
    
    ConfigurarGrdProductos
    CargarComboBoxes
    
    Select Case mi_Opcion
    Case sghAgregar
        Me.Caption = "Agregar Ventas"
    Case sghModificar
        Me.Caption = "Modificar Ventas"
    Case sghConsultar
        Me.Caption = "Consultar Ventas"
        
    Case sghEliminar
        Me.Caption = "Anular Ventas"
    End Select
    CargarDatosAlFormulario
    CargaItemsDebajoDeStockMinimo
    If mi_Opcion = sghAgregar Then
       'btnAceptar.Enabled = Not True     'licencia
    End If
End Sub
Sub ConfigurarGrdProductos()
    grdProductos.movNumero = ml_movNumero
    grdProductos.IdAlmacen = 0
    grdProductos.inicializar
    grdProductos.TipoPrecioParaNiNs = 3    'precio de venta

End Sub

Sub CargarDatosAlFormulario()
    mo_Formulario.HabilitarDeshabilitar Me.txtDocumento, False
    mo_Formulario.HabilitarDeshabilitar Me.txtNpreventa, False
    mo_Formulario.HabilitarDeshabilitar Me.txtTipoComprobante, False
    mo_Formulario.HabilitarDeshabilitar Me.txtFregistro, False
    mo_Formulario.HabilitarDeshabilitar Me.txtHoraRegistro, False
    mo_Formulario.HabilitarDeshabilitar Me.txtEstado, False
    mo_Formulario.HabilitarDeshabilitar Me.txtNombrePaciente, False
    mo_Formulario.HabilitarDeshabilitar Me.txtNombreDx, False
    mo_Formulario.HabilitarDeshabilitar Me.txtTurno, False
    mo_Formulario.HabilitarDeshabilitar Me.txtVendedor, False
    mo_Formulario.HabilitarDeshabilitar Me.txtCajero, False
    mo_Formulario.HabilitarDeshabilitar Me.txtCaja, False
    mo_Formulario.HabilitarDeshabilitar Me.txtDatosDeCuenta, False
    mo_Formulario.HabilitarDeshabilitar Me.txtPlan, False
    mo_Formulario.HabilitarDeshabilitar txtRedondeo, False
    wxParametro302 = lcBuscaParametro.SeleccionaFilaParametro(302)
    wxParametro509 = lcBuscaParametro.SeleccionaFilaParametro(509)
    If lcBuscaParametro.SeleccionaFilaParametro(249) = "S" Then
       optVtaSinPlan.Visible = True
    End If
    Select Case mi_Opcion
     Case sghAgregar
        txtFregistro.Text = lcBuscaParametro.RetornaFechaServidorSQL      'Format(Now, sighentidades.DevuelveHoraSoloFormato_HM)
        txtHoraRegistro.Text = lcBuscaParametro.RetornaHoraServidorSQL
        grdProductos.movNumero = ""
        grdProductos.LimpiarGrilla
        grdProductos.CargaProductosPorMovNumero
        'grdProductos.AgregaRegistro
        optVentas_Click 1
     Case sghModificar
        DeshabilitaCabecera
        CargarDatosALosControles
     Case sghConsultar
        DeshabilitaCabecera
        CargarDatosALosControles
        btnAceptar.Enabled = False
     Case sghEliminar
        DeshabilitaCabecera
        CargarDatosALosControles
 End Select
End Sub

Sub DeshabilitaCabecera()
    If ml_IdTipoVentaSeleccionada = 0 Then
        mo_Formulario.HabilitarDeshabilitar Me.cmbAlmOrigen, False
        fraTipoVenta.Enabled = False
        mo_Formulario.HabilitarDeshabilitar Me.txtNcuenta, False
        cmdBuscaCuentaPorApellidos.Enabled = False
        mo_Formulario.HabilitarDeshabilitar Me.txtNhistoria, False
        btnBuscarPaciente.Enabled = False
        mo_Formulario.HabilitarDeshabilitar Me.cmbTipoFinanciamiento, False
    Else
        mo_Formulario.HabilitarDeshabilitar Me.cmbAlmOrigen, False
        fraTipoVenta.Enabled = False
        mo_Formulario.HabilitarDeshabilitar Me.txtNcuenta, False
        cmdBuscaCuentaPorApellidos.Enabled = False
        mo_Formulario.HabilitarDeshabilitar Me.txtNhistoria, False
        btnBuscarPaciente.Enabled = False
        mo_Formulario.HabilitarDeshabilitar Me.cmbTipoFinanciamiento, False
    End If
End Sub



Sub CargarComboBoxes()
    Dim rsIdAlmacen As Recordset
    '
    Set oRsItemsUnidosis = mo_ReglasFarmacia.farmUnidosisSeleccionarTodos
    '
    Set rsIdAlmacen = mo_AdminServiciosComunes.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghAlmacenFarmacia, ml_idUsuario)
    '
    mo_cmbAlmacenOrigen.BoundColumn = "IdAlmacen"
    mo_cmbAlmacenOrigen.ListField = "Descripcion"
    'SCCQ 03/06/2020 Cambio23 Inicio
     If rsIdAlmacen.RecordCount > 0 Then 'Solo filtra farmacias asignadas
      Set mo_cmbAlmacenOrigen.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idtipoLocales='F' and idTipoSuministro='01' and idEstado=1 and idAlmacen in (select idLaboraSubArea from EmpleadosLugarDeTrabajo where idLaboraArea=" + CStr(sghAlmacenFarmacia) + " and idEmpleado=" + CStr(ml_idUsuario) + ")")
     Else ' Muestra todas las farmacias como lo hacía antes
      Set mo_cmbAlmacenOrigen.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idtipoLocales='F' and idTipoSuministro='01' and idEstado=1")
     End If
    'SCCQ 03/06/2020 Cambio23 Fin
    ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
    'SCCQ 03/06/2020 Cambio23 Inicio
    If rsIdAlmacen.RecordCount = 1 Then
       mo_cmbAlmacenOrigen.BoundText = rsIdAlmacen.Fields!idLaboraSubArea
       'mo_Formulario.HabilitarDeshabilitar Me.cmbAlmOrigen, False
       lbLaFarmaciaEsUnidosis = mo_ReglasFarmacia.FarmaciaEsUnidosis(Val(mo_cmbAlmacenOrigen.BoundText))
    End If
    'SCCQ 03/06/2020 Cambio23 Fin
    'UNIDOSIS
    Dim lcMensajeLicencia As String
'    If False Then     'licencia
'       cmbFarmaciaOrigen.Visible = False
'       txtCodigoDespacho.Visible = False
'    Else
        Dim oRsTmpA As New Recordset
        Set oRsTmpA = mo_ReglasComunes.FuentesFinanciamientoSegunFiltro("idFuenteFinanciamiento=3")
        If oRsTmpA.RecordCount > 0 Then
           If Not IsNull(oRsTmpA!CuentaParaUnidosis) Then
              txtCodigoDespacho.Text = oRsTmpA!CuentaParaUnidosis
           End If
        End If
        oRsTmpA.Close
        Set oRsTmpA = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idtipoLocales='F' and idTipoSuministro='01' and idEstado=1 and esUnidosis=1")
        mo_cmbUnidosis.BoundColumn = "IdAlmacen"
        mo_cmbUnidosis.ListField = "Descripcion"
       ' Set mo_cmbUnidosis.RowSource = oRsTmpA
        If oRsTmpA.RecordCount = 1 Then
           oRsTmpA.MoveFirst
           mo_cmbUnidosis.BoundText = Trim(Str(oRsTmpA!IdAlmacen))
        End If
        Set oRsTmpA = Nothing
        Set mo_cmbFarmaciaOrigen.MiComboBox = cmbFarmaciaOrigen
        mo_cmbFarmaciaOrigen.BoundColumn = "IdAlmacen"
        mo_cmbFarmaciaOrigen.ListField = "Descripcion"
        Set mo_cmbFarmaciaOrigen.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idtipoLocales='F' and idTipoSuministro='01' and idEstado=1 and esUnidosis=0")
        ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
'        If rsIdAlmacen.RecordCount > 0 Then
'           mo_cmbFarmaciaOrigen.BoundText = rsIdAlmacen.Fields!idLaboraSubArea
'           mo_Formulario.HabilitarDeshabilitar Me.cmbFarmaciaOrigen, False
'        End If
'        mo_cmbFarmaciaOrigen.BoundText = lcBuscaParametro.SeleccionaFilaParametro(577)
        mo_Formulario.HabilitarDeshabilitar Me.cmbFarmaciaOrigen, False
 '   End If
    
    '
    mo_cmbPrescriptor.BoundColumn = "IdMedico"
    mo_cmbPrescriptor.ListField = "Dmedico"
    Set mo_cmbPrescriptor.RowSource = mo_ReglasDeProgMedica.MedicosSeleccionarTodosOrdenadoAlfabeticamente
    ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
    '
    mo_cmbTipoReceta.BoundColumn = "idTipoReceta"
    mo_cmbTipoReceta.ListField = "TipoReceta"
    Set mo_cmbTipoReceta.RowSource = mo_ReglasFarmacia.FarmTipoRecetaDevuelveTodos
    ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
    '
    Set oRsTipoFinanciamiento = mo_ReglasFarmacia.TipoFinanciamientosDevuelveSoloFarmacia("")
    mo_cmbTipoFinanciamiento.BoundColumn = "idTipoFinanciamiento"
    mo_cmbTipoFinanciamiento.ListField = "Descripcion"
    Set mo_cmbTipoFinanciamiento.RowSource = mo_ReglasFarmacia.TipoFinanciamientosDevuelveSoloFarmacia("")
    ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
    '
    If ms_MensajeError <> "" Then
       MsgBox ms_MensajeError
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
   If SIGHEntidades.ParaAuditoria = "" Then
      LimpiarVariablesDeMemoria
   ElseIf MsgBox("Hubo cambios, desea salir de todas maneras ?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
      LimpiarVariablesDeMemoria
      SIGHEntidades.ParaAuditoriaPorCadaDato sghAudLimpiar, ""
   End If
End Sub







Private Sub grdHistorico_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
On Error Resume Next
grdHistorico.Bands(0).Columns("nombre").Width = 6000
grdHistorico.Bands(0).Columns("codigo").Width = 700
grdHistorico.Bands(0).Columns("fechaCreacion").Width = 1000
grdHistorico.Bands(0).Columns("horaCreacion").Width = 800
grdHistorico.Bands(0).Columns("tipo").Width = 500
grdHistorico.Bands(0).Columns("movNumero").Width = 1100
End Sub



Private Sub grdProductos_SeIngresoProducto(lcCodigo As String)
    txtHtotCantidad.Text = ""
    If Val(lcCodigo) > 0 And Me.chkHistorico.Value = 1 And Me.optVentas.Value = True Then
       Dim oRsHistoricos As New Recordset
       Set oRsHistoricos = grdHistorico.DataSource
       oRsHistoricos.Filter = ""
       If oRsHistoricos.RecordCount > 0 Then
          If mi_Opcion = sghAgregar Then
             oRsHistoricos.Filter = "codigo='" & Trim(lcCodigo) & "'"
          Else
             oRsHistoricos.Filter = "codigo='" & Trim(lcCodigo) & "' and movNumero<>'" & Me.txtDocumento.Text & "'"
          End If
          If oRsHistoricos.RecordCount > 0 Then
             Dim lnCantidad As Long
             lnCantidad = 0
             oRsHistoricos.MoveFirst
             Do While Not oRsHistoricos.EOF
                lnCantidad = lnCantidad + oRsHistoricos!Cantidad
                oRsHistoricos.MoveNext
             Loop
             oRsHistoricos.MoveFirst
             txtHtotCantidad.Text = lnCantidad
          End If
       End If
       Set oRsHistoricos = Nothing
    End If
End Sub

Private Sub grdProductos_SePresionoTeclaEspecial(KeyCode As Integer)
     If KeyCode = vbKeyF2 Or KeyCode = vbKeyF3 Then
        AdministrarKeyPreview KeyCode
        'Me.KeyPreview = False
     End If
End Sub

Private Sub grdProductos_Totalizado(lnTotalIngresado As Double)
    If Me.optPreventa.Value = True Then
        txtRedondeo.Text = SIGHEntidades.DevuelveNumeroRedondeado(lnTotalIngresado)  'debb-mayo2014
    Else
        txtRedondeo.Text = lnTotalIngresado
    End If
End Sub



Private Sub optPreventa_Click(Value As Integer)
    If optPreventa.Value Then
    'SCCQ 23/10/2020 Cambio34 Inicio
    txtNcuenta.Text = ""
    txtDatosDeCuenta.Text = ""
    txtPlan.Text = ""
    txtNhistoria.Text = ""
    txtNombrePaciente.Text = ""
    'SCCQ 23/10/2020 Cambio34 Fin
        BlanquedaVariablesUnidosis
        lnEpsPorcentaje = 0
        Set oRsTipoFinanciamiento = mo_ReglasFarmacia.TipoFinanciamientosDevuelveSoloFarmacia(" and dbo.TiposFinanciamiento.TipoVenta='P'")
        lcPosicionDefaultCombo = ""
        If oRsTipoFinanciamiento.RecordCount = 1 Then
            lcPosicionDefaultCombo = Trim(Str(oRsTipoFinanciamiento.Fields!IdTipoFinanciamiento))
        End If
        mo_cmbTipoFinanciamiento.BoundColumn = "idTipoFinanciamiento"
        mo_cmbTipoFinanciamiento.ListField = "Descripcion"
        Set mo_cmbTipoFinanciamiento.RowSource = mo_ReglasFarmacia.TipoFinanciamientosDevuelveSoloFarmacia(" and dbo.TiposFinanciamiento.TipoVenta='P'")
        'If lcPosicionDefaultCombo <> "" Then
        '   mo_cmbTipoFinanciamiento.BoundText = lcPosicionDefaultCombo
        'End If
        mo_cmbTipoFinanciamiento.BoundText = "1"
        ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
        grdProductos.TipoVentaSeleccionada = 1
        mo_Formulario.HabilitarDeshabilitar Me.cmbTipoFinanciamiento, True
        mo_Formulario.HabilitarDeshabilitar Me.txtNcuenta, True
        mo_Formulario.HabilitarDeshabilitar Me.txtObservaciones, False
        If mi_Opcion = sghAgregar Then
           cmbTipoReceta.Text = ""
           cmbTipoReceta.SetFocus
        End If
        FraRedondeo.Visible = True
        chkPlanNoCubre.Enabled = False
        lnIdTipoServicio = 0
        btnBuscarPaciente.Enabled = False
        cmdBuscaCuentaPorApellidos.Enabled = True
        HabilitaDatosPacientePreVenta True
    End If
End Sub

Sub HabilitaDatosPacientePreVenta(lValue As Boolean)
    mo_Formulario.HabilitarDeshabilitar txtNhistoria, lValue
    mo_Formulario.HabilitarDeshabilitar txtNombrePaciente, lValue
    txtNhistoria.ToolTipText = IIf(lValue = True, "N° DNI", "N° Historia")
    Label6.Caption = IIf(lValue = True, "N° DNI", "Paciente")
End Sub

Sub BlanquedaVariablesUnidosis()
    lnCuentaUnidosis = 0
    lbLaFuenteFinanciamientoUsadoEnFUnidosis = False
    cmbUnidosis.Visible = False
    lblUnidosis.Visible = False
End Sub

Private Sub optVentas_Click(Value As Integer)
    If optVentas.Value Then
        BlanquedaVariablesUnidosis
        Set oRsTipoFinanciamiento = mo_ReglasFarmacia.TipoFinanciamientosDevuelveSoloFarmacia(" and dbo.TiposFinanciamiento.esFuenteFinanciamiento=1")
        mo_cmbTipoFinanciamiento.BoundColumn = "idTipoFinanciamiento"
        mo_cmbTipoFinanciamiento.ListField = "Descripcion"
        Set mo_cmbTipoFinanciamiento.RowSource = mo_ReglasFarmacia.TipoFinanciamientosDevuelveSoloFarmacia(" and dbo.TiposFinanciamiento.esFuenteFinanciamiento=1")
        grdProductos.TipoVentaSeleccionada = 0
        ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
        mo_Formulario.HabilitarDeshabilitar Me.cmbTipoFinanciamiento, False
        mo_Formulario.HabilitarDeshabilitar Me.txtNhistoria, False
        mo_Formulario.HabilitarDeshabilitar Me.txtObservaciones, True
        mo_Formulario.HabilitarDeshabilitar Me.txtNcuenta, True
        If mi_Opcion = sghAgregar Then
           mo_cmbTipoReceta.BoundText = "1"
        End If
        FraRedondeo.Visible = False
        chkPlanNoCubre.Enabled = True
        lnIdTipoServicio = 0
        btnBuscarPaciente.Enabled = False
        cmdBuscaCuentaPorApellidos.Enabled = True
        HabilitaDatosPacientePreVenta False
    End If
End Sub

Private Sub optVtaSinPlan_Click(Value As Integer)
    If optVtaSinPlan.Value Then
        BlanquedaVariablesUnidosis
        lnEpsPorcentaje = 0
        Me.grdProductos.LimpiarGrilla
        Me.txtNcuenta.Text = ""
        Me.txtNreceta.Text = ""
        Me.txtDatosDeCuenta.Text = ""
        Me.txtNombrePaciente.Text = ""
        Me.txtNhistoria.Text = ""
        Set oRsTipoFinanciamiento = mo_ReglasFarmacia.TipoFinanciamientosDevuelveSoloFarmacia(" and dbo.TiposFinanciamiento.idTipoFinanciamiento=6 or dbo.TiposFinanciamiento.idTipoFinanciamiento=9 or dbo.TiposFinanciamiento.idTipoFinanciamiento=10")
        mo_cmbTipoFinanciamiento.BoundColumn = "idTipoFinanciamiento"
        mo_cmbTipoFinanciamiento.ListField = "Descripcion"
        Set mo_cmbTipoFinanciamiento.RowSource = mo_ReglasFarmacia.TipoFinanciamientosDevuelveSoloFarmacia(" and dbo.TiposFinanciamiento.idTipoFinanciamiento=6 or dbo.TiposFinanciamiento.idTipoFinanciamiento=9 or dbo.TiposFinanciamiento.idTipoFinanciamiento=10")
        grdProductos.TipoVentaSeleccionada = 0
        ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
        mo_Formulario.HabilitarDeshabilitar Me.txtNcuenta, False
        mo_Formulario.HabilitarDeshabilitar Me.txtNhistoria, False
        mo_Formulario.HabilitarDeshabilitar Me.txtObservaciones, True
        mo_Formulario.HabilitarDeshabilitar Me.cmbTipoFinanciamiento, True
        If mi_Opcion = sghAgregar Then
           cmbTipoReceta.Text = ""
           cmbTipoFinanciamiento.SetFocus
        End If
        FraRedondeo.Visible = False
        chkPlanNoCubre.Enabled = False
        lnIdTipoServicio = 0
        btnBuscarPaciente.Enabled = False
        cmdBuscaCuentaPorApellidos.Enabled = False
    End If
End Sub

Private Sub txtCodigoDespacho_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtCodigoDespacho
End Sub

Private Sub txtCodigoDespacho_LostFocus()
    txtNcuenta.Text = txtCodigoDespacho.Text
    txtNcuenta_LostFocus
End Sub

Private Sub txtDx_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDx

End Sub


Private Sub txtDx_LostFocus()
        Dim oDODiagnostico As DODiagnostico
        Set oDODiagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorCodigoCIE2004(txtDx.Text, True)
        If Not oDODiagnostico Is Nothing Then
            ml_IdDiagnostico = oDODiagnostico.idDiagnostico
            txtNombreDx.Text = oDODiagnostico.Descripcion
        Else
            ml_IdDiagnostico = 0
            txtNombreDx.Text = ""
        End If

End Sub

Private Sub txtFprescribe_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbPrescriptor
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtFprescribe_LostFocus()
If Not IsDate(txtFprescribe.Text) Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        txtFprescribe.Text = SIGHEntidades.FECHA_VACIA_DMY_HM
        Exit Sub
    End If
End Sub

Private Sub txtNcuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNcuenta

End Sub


Private Sub txtNcuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    ElseIf Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
    End If

End Sub

Private Sub txtNcuenta_LostFocus()
   'BlanquedaVariablesUnidosis
   grdProductos.PermiteAgregarItems = True
   If mo_Teclado.TextoEsSoloNumeros(txtNcuenta.Text) Then
       Dim oRsTmp As New Recordset
       Dim lbSigue As Boolean, lnTotal1 As Double
       Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
       Dim oConexion As New Connection
       oConexion.Open SIGHEntidades.CadenaConexion
       oConexion.CursorLocation = adUseClient
       Set oRsTmp = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(txtNcuenta.Text, oConexion)
       txtDatosDeCuenta.Text = ""
       If mi_Opcion = sghAgregar Then
          cmbTipoFinanciamiento.Text = ""
          txtTipoComprobante.Text = ""
       End If
       ml_IdPaciente = 0
       ml_idFuenteFinanciamiento = 0
       txtNombrePaciente.Text = ""
       txtNhistoria.Text = ""
       
       txtPlan.Text = ""
       lbSigue = True
       If oRsTmp.RecordCount > 0 Then
          If oRsTmp.Fields!idEstado <> 1 And lbCuentaDeEmergenciaCerrada = False Then
             If mi_Opcion <> sghConsultar Then
                MsgBox "Ese estado de Cuenta no se encuentra ABIERTA", vbInformation, Me.Caption
                If mi_Opcion = sghModificar Or mi_Opcion = sghEliminar Then
                   btnAceptar.Enabled = False
                Else
                   lbSigue = False
                End If
             End If
          End If
          '
          If lbSigue = True And oRsTmp!esPacienteExterno <> True And wxParametro509 = "S" And mi_Opcion = sghAgregar Then
             If Val(txtNreceta.Text) = 0 Then
                MsgBox "No puede usar N° CUENTA, tiene que generar RECETA", vbInformation, Me.Caption
                lbSigue = False
             End If
          End If
          'unidosis
'          lbLaFuenteFinanciamientoUsadoEnFUnidosis = mo_ReglasComunes.FuenteFinanciamientoEsUnidosis(oRsTmp!idFuenteFinanciamiento, lnCuentaUnidosis)
'          If lbSigue = True And mi_Opcion = sghAgregar Then
'             If lbLaFarmaciaEsUnidosis = True And lbLaFuenteFinanciamientoUsadoEnFUnidosis = False Then
'                MsgBox "La FUENTE FINANCIAMIENTO de la Cuenta, no se usa para despachar en FARMACIA UNIDOSIS", vbInformation, ""
'                lbSigue = False
'             End If
'             If lbSigue = True And lbLaFarmaciaEsUnidosis = True And lnCuentaUnidosis = Val(txtNcuenta.Text) Then
'                MsgBox "La Cuenta, solo se usa para despachar hacia la FARMACIA UNIDOSIS", vbInformation, ""
'                lbSigue = False
'             End If
'             If lbSigue = True And lbLaFarmaciaEsUnidosis = False And lnCuentaUnidosis = Val(txtNcuenta.Text) Then
'                cmbUnidosis.Visible = True
'                lblUnidosis.Visible = True
'             End If
'          End If
          '
          'Chequea si es Consulta Externa/Paciente particular
          If lbSigue = True And mi_Opcion = sghAgregar And _
             oRsTmp.Fields!IdTipoServicio = sghTipoServicio.sghConsultaExterna And _
             mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(oRsTmp.Fields!IdFormaPago, oConexion) = True Then
               If MsgBox("Ese N° Cuenta es de CONSULTA EXTERNA de un Paciente PAGANTE" & Chr(13) & Chr(13) & _
                         "Está seguro de despachar Medicamentos ?", vbQuestion + vbYesNo, "Mensaje") = vbNo Then
                  lbSigue = False
               Else
                  Me.optPreventa.Value = True
               End If
          ElseIf Me.optPreventa.Value = True And lbSigue = True And mi_Opcion = sghAgregar And mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(oRsTmp.Fields!IdFormaPago, oConexion) = False Then
              Me.optVentas.Value = True
          End If
          If mi_Opcion = sghAgregar And _
             mo_AdminAdmision.AtencionesDatosAdicionalesSItieneCodigoPrestacionSIS(Val(txtNcuenta.Text), wxParametro302, _
                                                                          oRsTmp.Fields!idFuenteFinanciamiento) = False Then
             lbSigue = False
          End If
          If mi_Opcion = sghAgregar And _
                                    mo_AdminAdmision.LaFechaDespachoEsMenorAfechaCita(CDate(Format(oRsTmp!fechaingreso, _
                                    SIGHEntidades.DevuelveFechaSoloFormato_DMY) & " " & oRsTmp!horaIngreso)) = True Then
             lbSigue = False
          End If
          If lbSigue Then
                lnEpsPorcentaje = mo_ReporteUtil.DevuelveEpsPorcentaje(oRsTmp!EpsPorcentaje)
                mo_Formulario.HabilitarAlerta txtPlan, IIf(lnEpsPorcentaje > 0, True, False)
                If lnEpsPorcentaje > 0 And mi_Opcion <> sghAgregar Then
                   '
                   Dim lcBoletaEPS As String
                   lblOrdenPago.Tag = mo_ReglasFacturacion.DevuelveOrdenPago(oRsTmp!idAtencion, sghPtoCargaFarmacia, mo_DoFarmMovimiento.fechaCreacion, oConexion, lcBoletaEPS)
                   lblOrdenPago.Caption = "N° Orden de Pago: " & lblOrdenPago.Tag
                   If lcBoletaEPS <> "" Then
                        lblOrdenPago.Caption = lcBoletaEPS
                        MsgBox "El SEGURO tiene EPS, No podrá MODIFICAR/ELIMINAR porque ya pagó en CAJA" & Chr(13) & _
                               "Tendría que ANULAR (o NOTA DE CREDITO) la BOLETA para usar MODIFICAR/ELIMINAR", vbInformation, ""
                        Me.btnAceptar.Enabled = False
                   End If
                   '
                End If
                
                lnIdTipoServicio = oRsTmp.Fields!IdTipoServicio
                txtDatosDeCuenta.Text = "F.Ing: " & oRsTmp.Fields!fechaingreso & _
                                       " (" & IIf(oRsTmp!esPacienteExterno = True, "Ext", _
                                              IIf(oRsTmp.Fields!IdTipoServicio = 1, "CExt", _
                                              IIf(oRsTmp.Fields!IdTipoServicio = 3, "Hosp", "Emer"))) & _
                                       ") (Est: " & Trim(oRsTmp.Fields!estadoCta) & ")"   'debb-jamo
                txtPlan.Text = "IAFA Act.: " & oRsTmp.Fields!dFuenteFinanciamiento & mo_ReporteUtil.DevuelveEPScubre(lnEpsPorcentaje)
                ml_IdPaciente = oRsTmp.Fields!IdPaciente
                ml_idFuenteFinanciamiento = oRsTmp.Fields!idFuenteFinanciamiento
                
                txtNombrePaciente.Text = Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & " " & oRsTmp.Fields!PrimerNombre
                txtNhistoria.Text = HCigualDNI_DevuelveHistoriaConCerosIzquierda(oRsTmp.Fields!NroHistoriaClinica, False)
                ml_idTipoConcepto = mo_ReglasFacturacion.FuentesFinanciamientoDevuelveIdTipoConceptoFarmacia(ml_idFuenteFinanciamiento, oConexion)
                If mi_Opcion = sghAgregar Then
          
                      '
                      txtDatosDeCuenta = txtDatosDeCuenta & " (" & Trim(mo_ReglasFacturacion.BuscaServicioActualDelPaciente(mo_ReglasFarmacia.DevuelveServicioDondeSeEncuentraElPacienteSegunFechaHora(Val(txtNcuenta.Text), CDate(Me.txtFregistro.Text), lcBuscaParametro.RetornaHoraServidorSQL))) & ")"   'debb-jamo
                      If optVentas.Value = True Then
                         mo_cmbTipoFinanciamiento.BoundText = oRsTmp.Fields!IdFormaPago
                      Else
                         cmbTipoFinanciamiento_Click
                         mo_cmbTipoFinanciamiento.BoundText = "1"
                         mo_cmbTipoReceta.BoundText = "1"
                      End If
                      '
                      If lnIdTipoServicio <> sghTipoServicio.sghHospitalizacion And cmbPrescriptor.Text = "" Then
                         mo_cmbPrescriptor.BoundText = oRsTmp.Fields!idMedicoIngreso
                         If lnIdTipoServicio = sghTipoServicio.sghConsultaExterna And Not IsNull(oRsTmp.Fields!fechaEgreso) Then
                            txtFprescribe.Text = Format(oRsTmp.Fields!fechaEgreso & " " & oRsTmp.Fields!horaEgreso, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HM)
                         End If
                      End If
                      '
                      Set oRsTmp = mo_AdminAdmision.AtencionesDiagnosticosSeleccionarTodosPorIdAtencion(oRsTmp.Fields!idAtencion)
                      If oRsTmp.RecordCount > 0 Then
                           txtDx.Text = oRsTmp.Fields!CodigoCIE2004
                           txtNombreDx.Text = oRsTmp.Fields!Descripcion
                           ml_IdDiagnostico = oRsTmp.Fields!idDiagnostico
                      ElseIf wxParametro578 = "S" Then
                           MsgBox "No se puede despachar sino tiene Dx", vbInformation, ""
                           txtDatosDeCuenta.Text = ""
                           txtNombrePaciente.Text = ""
                           txtPlan.Text = ""
                           mo_cmbTipoFinanciamiento.BoundText = ""
                           mo_cmbPrescriptor.BoundText = ""
                      End If
                      '
                      MuestraHistoricoDelPacienteXcuenta    'kike 2017
                      '
                      grdProductos.TabEnDescripcion
                Else
                      If Me.txtFregistro.Text <> SIGHEntidades.FECHA_VACIA_DMY Then
                         txtDatosDeCuenta = txtDatosDeCuenta & " (" & Trim(mo_ReglasFacturacion.BuscaServicioActualDelPaciente(mo_ReglasFarmacia.DevuelveServicioDondeSeEncuentraElPacienteSegunFechaHora(Val(txtNcuenta.Text), CDate(Me.txtFregistro.Text), Me.txtHoraRegistro.Text))) & ")"   'debb-jamo
                      End If
                      If ml_idFuenteFinanciamiento <> ml_IdFuenteFinanciamientoDespacho And ml_IdFuenteFinanciamientoDespacho > 0 Then
                         txtPlan.Text = "Plan Desp: " & Trim(mo_ReglasFacturacion.FuentesFinanciamientoDevuelveNombrePlan(ml_IdFuenteFinanciamientoDespacho)) & " - " & txtPlan.Text
                      End If
                End If
          Else
                txtNreceta.Text = ""
                
          End If
       End If
       oRsTmp.Close
       Set oRsTmp = Nothing
       oConexion.Close
       Set oConexion = Nothing
   End If
End Sub

Private Sub txtNhistoria_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNhistoria

End Sub


Private Sub btnBuscarPaciente_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaPacientes
    Dim oDOPaciente As New DOPaciente
    Dim oConexion As New Connection
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oBusqueda.TipoFiltro = sghFiltrarTodos
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(oBusqueda.IdRegistroSeleccionado, oConexion)
        If Not oDOPaciente Is Nothing Then
            ml_IdPaciente = oDOPaciente.IdPaciente
            txtNhistoria.Text = oDOPaciente.NroHistoriaClinica
            txtNombrePaciente.Text = Trim(oDOPaciente.ApellidoPaterno) + " " + Trim(oDOPaciente.ApellidoMaterno) + " " + oDOPaciente.PrimerNombre
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
End Sub
Private Sub cmdBuscaDx_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaDiagnosticos
    Dim oDODiagnostico As DODiagnostico
    oBusqueda.SoloMuestraDxGalenHos = True
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDODiagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
        If Not oDODiagnostico Is Nothing Then
            ml_IdDiagnostico = oDODiagnostico.idDiagnostico
            txtDx.Text = oDODiagnostico.CodigoCIE2004
            txtNombreDx.Text = oDODiagnostico.Descripcion
        End If
    End If
End Sub


Sub CargaDatosVendedorCajero(lnIdVendedorCajero As Long, EsVendedor As Boolean)
    Dim oDOEmpleado As dOEmpleado
    Set oDOEmpleado = mo_AdminServiciosComunes.EmpleadosSeleccionarPorId(lnIdVendedorCajero)
    With oDOEmpleado
        If EsVendedor Then
           txtVendedor.Text = Trim(.ApellidoPaterno) & " " & Trim(.ApellidoMaterno) & .Nombres
        Else
           txtCajero.Text = Trim(.ApellidoPaterno) & " " & Trim(.ApellidoMaterno) & .Nombres
        End If
    End With
End Sub

Sub CargarDatosALosControles()
   btnNuevo.Visible = False
   If ml_IdTipoVentaSeleccionada = 0 Then
        CargaVentasDirectas
        '******permiso a Modificar documento con Fecha Anterior a la actual
        Dim mo_PermisosFacturacion As New PermisosFacturacion
        Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
        Set mo_PermisosFacturacion = mo_ReglasSeguridad.UsuariosRolesSeleccionarPermisosFacturacion(ml_idUsuario)
        If mo_PermisosFacturacion.ActualizaFechaDocumentoES = False Then
           If CDate(lcBuscaParametro.RetornaFechaServidorSQL) <> CDate(txtFregistro.Text) Then
              MsgBox "No tiene ACCESO a Modificar/Anular una Venta" & Chr(13) & " de una Fecha Registro diferente a la actual", vbExclamation, Me.Caption
              btnAceptar.Enabled = False
           End If
        End If
        Set mo_PermisosFacturacion = Nothing
        Set mo_ReglasSeguridad = Nothing
        If mi_Opcion = sghConsultar Then
           btnImprimir.Visible = True
        End If
   Else
        CargaPreVenta
   End If
   DeshabilitaCabecera
   If lnIdReceta > 0 Then
      grdProductos.PermiteAgregarItems = False
   End If
   If txtNcuenta.Text = "" And Me.optPreventa.Value = True Then
        HabilitaDatosPacientePreVenta True
   End If
End Sub
Sub CargaPreVenta()
    Dim oConexion As New Connection
    Dim oRsTmp As New Recordset
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    mo_DofarmPreVenta.idPreVenta = Val(ml_movNumero)
    If Not mo_ReglasFarmacia.FarmPreventasSeleccionarPorId(mo_DofarmPreVenta) Then
      MsgBox mo_ReglasFarmacia.MensajeError
      Exit Sub
    End If
    txtNpreventa.Text = ml_movNumero
    mo_cmbAlmacenOrigen.BoundText = mo_DofarmPreVenta.IdAlmacen
    txtNcuenta.Text = IIf(mo_DofarmPreVenta.idCuentaAtencion > 0, mo_DofarmPreVenta.idCuentaAtencion, "")
    txtNcuenta_LostFocus
    optVentas.Value = False: optPreventa.Value = True
    mo_cmbTipoFinanciamiento.BoundText = mo_DofarmPreVenta.IdTipoFinanciamiento
    mo_cmbTipoReceta.BoundText = mo_DofarmPreVenta.idTipoReceta
    ml_IdVendedor = mo_DofarmPreVenta.idVendedor
    CargaDatosVendedorCajero ml_IdVendedor, True
    mo_cmbPrescriptor.BoundText = mo_DofarmPreVenta.idPrescriptor
    txtFprescribe.Text = Format(mo_DofarmPreVenta.FechaHoraPrescribe, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HM)
    'Paciente
    If txtNcuenta.Text = "" Then
        ml_IdPaciente = mo_DofarmPreVenta.IdPaciente
        If ml_IdPaciente > 0 Then
            mo_DoPaciente.IdPaciente = ml_IdPaciente
            Set mo_DoPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(ml_IdPaciente, oConexion)
            txtNhistoria.Text = mo_DoPaciente.nroDocumento  ' mo_DoPaciente.NroHistoriaClinica
            txtNombrePaciente.Text = Trim(mo_DoPaciente.ApellidoPaterno) & " " & Trim(mo_DoPaciente.ApellidoMaterno) & " " & mo_DoPaciente.PrimerNombre
        Else
            txtNhistoria.Text = mo_DofarmPreVenta.dni
            txtNombrePaciente.Text = mo_DofarmPreVenta.paciente
        End If
    End If
    'Dx
    Dim mo_Diagnostico As New DODiagnostico
    ml_IdDiagnostico = mo_DofarmPreVenta.idDiagnostico
    If ml_IdDiagnostico > 0 Then
        Set mo_Diagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorId(ml_IdDiagnostico)
        txtDx.Text = mo_Diagnostico.CodigoCIE2004
        txtNombreDx.Text = mo_Diagnostico.Descripcion
    End If
   '
   Set oRsTmp = mo_ReglasComunes.RecetaCabeceraFiltraXcuentaYDocumentodespacho(Trim(txtNpreventa.Text), Val(txtNcuenta.Text))
   ms_MensajeError = ""
   lnIdReceta = 0
   If oRsTmp.RecordCount > 0 Then
       lnIdReceta = oRsTmp.Fields!idReceta
   End If
   oRsTmp.Close
   '
   '**************Datos de la tabla FarmMovimientoDetalle *****************
   grdProductos.idPreVenta = Val(ml_movNumero)
   grdProductos.TipoVentaSeleccionada = 1   'Preventas
   grdProductos.CargaProductosPorIdPreVenta
   grdProductos.RefrescarDatos
   lnTotalDocumento = grdProductos.DevuelveTotal
   txtEstado.Text = mo_ReglasFarmacia.DevuelveEstadoActualDePreventa("idEstadoPreventa=" & mo_DofarmPreVenta.idEstadoPreventa)
   If mo_DofarmPreVenta.idEstadoPreventa <> 1 Then
      btnAceptar.Enabled = False
   End If
   If mo_DofarmPreVenta.idEstadoPreventa = 2 Or mo_DofarmPreVenta.idEstadoPreventa = 0 Then
      Dim oDoComprobantePago As New DOCajaComprobantesPago
      Set oRsTmp = mo_ReglasFarmacia.FarmMovimientoVentasSeleccionarPorIdPreventa(mo_DofarmPreVenta.idPreVenta)
      If oRsTmp.RecordCount > 0 Then
         txtDocumento.Text = oRsTmp.Fields!DocumentoNumero
      End If
      oRsTmp.Close
      'Set oRsTmp = mo_ReglasCaja.CajaComprobantePagoSeleccionarPorFechaOdocumento(Left(txtDocumento.Text, 3), Mid(txtDocumento.Text, 5, 15), Date, Date)
      Set oRsTmp = mo_ReglasCaja.CajaComprobantePagoSeleccionarPorFechaOdocumento(Left(txtDocumento.Text, InStr(txtDocumento.Text, "-") - 1), Mid(txtDocumento.Text, InStr(txtDocumento.Text, "-") + 1, 15), Date, Date)
      If oRsTmp.RecordCount > 0 Then
         txtCaja.Text = IIf(IsNull(oRsTmp.Fields!dCaja), "", oRsTmp.Fields!dCaja)
         txtCajero.Text = IIf(IsNull(oRsTmp.Fields!ApellidoPaterno), "", Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & " " & oRsTmp.Fields!Nombres)
         txtTurno.Text = IIf(IsNull(oRsTmp.Fields!Descripcion), "", oRsTmp.Fields!Descripcion)
      End If
      oRsTmp.Close
   End If
   If lnTotalDocumento <> mo_DofarmPreVenta.Total Then
      FraRedondeo.Visible = True
      txtRedondeo.Text = mo_DofarmPreVenta.Total
   End If
   oConexion.Close
   'debb-29/12/2016
   If lnIdReceta > 0 Then
       grdProductos.esReceta = True
       Set oRsTmp = mo_ReglasComunes.RecetaCabeceraDetalleSeleccionaPorNroReceta(lnIdReceta)
       grdProductos.CargaCantidadRecetada oRsTmp
   End If
   
   Set oConexion = Nothing
   Set oRsTmp = Nothing

End Sub
Sub CargaVentasDirectas()
   Dim oConexion As New Connection
   Dim oRsTmp As New Recordset
   oConexion.Open SIGHEntidades.CadenaConexion
   oConexion.CursorLocation = adUseClient
 '**************Datos de la tabla FarmMovimiento *****************
   chkPlanNoCubre.Visible = False: txtDatosDeCuenta.Width = txtDatosDeCuenta.Width + chkPlanNoCubre.Width + 110
   mo_DoFarmMovimiento.movNumero = ml_movNumero
   mo_DoFarmMovimiento.MovTipo = lcConstanteMovimientoSalida
   If Not mo_ReglasFarmacia.FarmMovimientoSeleccionarPorId(mo_DoFarmMovimiento) Then
      MsgBox mo_ReglasFarmacia.MensajeError
      Exit Sub
   End If
   '
   txtDocumento.Text = mo_DoFarmMovimiento.DocumentoNumero
   mo_cmbAlmacenOrigen.BoundText = mo_DoFarmMovimiento.IdAlmacenOrigen
   txtObservaciones.Text = mo_DoFarmMovimiento.Observaciones
   txtEstado.Text = mo_ReglasFarmacia.DevuelveEstadoActualDelMovimiento("idEstadoMovimiento=" & mo_DoFarmMovimiento.idEstadoMovimiento)
   txtFregistro.Text = Format(mo_DoFarmMovimiento.fechaCreacion, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
   txtHoraRegistro.Text = Format(mo_DoFarmMovimiento.fechaCreacion, SIGHEntidades.DevuelveHoraSoloFormato_HM)

   '**************Datos de la tabla FarmMovimientoVentas *****************
   
   Dim mo_DoPaciente As New DOPaciente
   Dim mo_Diagnostico As New DODiagnostico
   With mo_DoFarmMovimientoVentas
       .movNumero = ml_movNumero
       .MovTipo = lcConstanteMovimientoSalida
       Me.txtFprescribe.Text = Format(.FechaHoraPrescribe, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HM)
       If Not mo_ReglasFarmacia.farmMovimientoVentasSeleccionarPorId(mo_DoFarmMovimientoVentas) Then
            MsgBox mo_ReglasFarmacia.MensajeError
            Exit Sub
       Else
            txtNpreventa.Text = .idPreVenta
            mo_cmbPrescriptor.BoundText = .idPrescriptor
            txtFprescribe.Text = Format(.FechaHoraPrescribe, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HM)
            txtNcuenta.Text = .idCuentaAtencion
            txtNcuenta_LostFocus
            optPreventa.Value = False: optVentas.Value = True
            mo_cmbTipoFinanciamiento.BoundText = .IdTipoFinanciamiento
            mo_cmbTipoReceta.BoundText = .idTipoReceta
            mo_cmbPrescriptor.BoundText = .idPrescriptor
            ml_IdFuenteFinanciamientoDespacho = .idFuenteFinanciamiento
            'debb-14/04/2011
            If mo_Teclado.TextoEsSoloNumeros(txtNcuenta.Text) Then
                Dim oRsTmp1 As New Recordset
                Set oRsTmp1 = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(Val(txtNcuenta.Text), oConexion)
                If mi_Opcion = sghModificar And oRsTmp1.Fields!idFuenteFinanciamiento <> ml_IdFuenteFinanciamientoDespacho Then
                   MsgBox "No se podrá modificar datos, porque el despacho tubo otra PRODUCTO/PLAN" & Chr(13) & "hubo RECALCULO", vbInformation, Me.Caption
                   btnAceptar.Enabled = False
                End If
                Set oRsTmp1 = Nothing
            End If
            'Dx
            ml_IdDiagnostico = .idDiagnostico
            If ml_IdDiagnostico > 0 Then
                Set mo_Diagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorId(.idDiagnostico)
                txtDx.Text = mo_Diagnostico.CodigoCIE2004
                txtNombreDx.Text = mo_Diagnostico.Descripcion
                Set mo_Diagnostico = Nothing
            End If
       End If
   End With
   '
   Set oRsTmp = mo_ReglasComunes.RecetaCabeceraFiltraXcuentaYDocumentodespacho(Trim(mo_DoFarmMovimiento.DocumentoNumero), Val(txtNcuenta.Text))
   lnIdReceta = 0
   If oRsTmp.RecordCount > 0 Then
       lnIdReceta = oRsTmp.Fields!idReceta
   End If
   oRsTmp.Close
   '
   CargaDatosVendedorCajero mo_DoFarmMovimiento.idUsuario, True
   If Val(txtNpreventa.Text) > 0 Then
      Dim oDoComprobantePago As New DOCajaComprobantesPago
      'Set oRsTmp = mo_ReglasCaja.CajaComprobantePagoSeleccionarPorFechaOdocumento(Left(txtDocumento.Text, 3), Mid(txtDocumento.Text, 5, 15), Date, Date)
      Set oRsTmp = mo_ReglasCaja.CajaComprobantePagoSeleccionarPorFechaOdocumento(Left(txtDocumento.Text, InStr(txtDocumento.Text, "-") - 1), Mid(txtDocumento.Text, InStr(txtDocumento.Text, "-") + 1, 15), Date, Date)
      If oRsTmp.RecordCount > 0 Then
         txtCaja.Text = IIf(IsNull(oRsTmp.Fields!dCaja), "", oRsTmp.Fields!dCaja)
         txtCajero.Text = IIf(IsNull(oRsTmp.Fields!ApellidoPaterno), "", Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & " " & oRsTmp.Fields!Nombres)
         txtTurno.Text = IIf(IsNull(oRsTmp.Fields!Descripcion), "", oRsTmp.Fields!Descripcion)
      End If
      oRsTmp.Close
      mo_DofarmPreVenta.idPreVenta = Val(txtNpreventa.Text)
      If Not mo_ReglasFarmacia.FarmPreventasSeleccionarPorId(mo_DofarmPreVenta) Then
         MsgBox mo_ReglasFarmacia.MensajeError
         Exit Sub
      End If
      If Val(txtNcuenta.Text) = 0 And mo_DofarmPreVenta.IdPaciente > 0 Then
         mo_DoPaciente.ApellidoPaterno = ""
         Set mo_DoPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(mo_DofarmPreVenta.IdPaciente, oConexion)
         If mo_DoPaciente.ApellidoPaterno <> "" Then
             ml_IdPaciente = mo_DofarmPreVenta.IdPaciente
             txtNombrePaciente.Text = Trim(mo_DoPaciente.ApellidoPaterno) & " " & Trim(mo_DoPaciente.ApellidoMaterno) & " " & mo_DoPaciente.PrimerNombre
             txtNhistoria.Text = mo_DoPaciente.NroHistoriaClinica
         End If
      End If
      'CargaDatosVendedorCajero mo_DofarmPreVenta.idVendedor, True
   End If
   If mo_cmbTipoFinanciamiento.BoundText = "6" Or mo_cmbTipoFinanciamiento.BoundText = "10" Then  'defensa nacional,credito personal
        mo_DoPaciente.ApellidoPaterno = ""
        Set mo_DoPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(mo_DoFarmMovimientoVentas.IdPaciente, oConexion)
        If mo_DoPaciente.ApellidoPaterno <> "" Then
            ml_IdPaciente = mo_DoFarmMovimientoVentas.IdPaciente
            txtNombrePaciente.Text = Trim(mo_DoPaciente.ApellidoPaterno) & " " & Trim(mo_DoPaciente.ApellidoMaterno) & " " & mo_DoPaciente.PrimerNombre
            txtNhistoria.Text = mo_DoPaciente.NroHistoriaClinica
        End If
   End If
   '**************Datos de la tabla FarmMovimientoDetalle *****************
   grdProductos.movNumero = ml_movNumero
   grdProductos.TipoVentaSeleccionada = 0   'VentaDirecta
   grdProductos.CargaProductosPorMovNumero
   grdProductos.RefrescarDatos
   lnTotalDocumento = grdProductos.DevuelveTotal
   'If mo_DoFarmMovimiento.IdEstadoMovimiento <> 1 Then
   '   btnAceptar.Enabled = False
   'End If
   If Val(txtNpreventa.Text) > 0 Then
      btnAceptar.Enabled = False
   End If
   If mo_DoFarmMovimiento.idEstadoMovimiento = 0 Then
      btnAceptar.Enabled = False
   End If
   If mo_cmbTipoFinanciamiento.BoundText = "1" And txtNpreventa.Text <> "" Then
        If lnTotalDocumento <> mo_DofarmPreVenta.Total Then
           FraRedondeo.Visible = True
           txtRedondeo.Text = mo_DoFarmMovimiento.Total
        End If
   End If
   If oRsTipoFinanciamiento.Fields!esFuenteFinanciamiento = False Then
       optVtaSinPlan.Value = True: optPreventa.Value = False: optVentas.Value = False
       mo_cmbTipoFinanciamiento.BoundText = mo_DoFarmMovimientoVentas.IdTipoFinanciamiento
   End If
   oConexion.Close
   'debb-29/12/2016
   If lnIdReceta > 0 Then
       grdProductos.esReceta = True
       Set oRsTmp = mo_ReglasComunes.RecetaCabeceraDetalleSeleccionaPorNroReceta(lnIdReceta)
       grdProductos.CargaCantidadRecetada oRsTmp
   End If
   Set oConexion = Nothing
   Set oRsTmp = Nothing

   'unidosis
   If Trim(mo_DoFarmMovimiento.Observaciones) <> "" Then
        mo_farmMovimiento1.movNumero = Trim(mo_DoFarmMovimiento.Observaciones)
        mo_farmMovimiento1.MovTipo = lcConstanteMovimientoEntrada
        mo_farmMovimiento1.IdUsuarioAuditoria = mo_DoFarmMovimiento.IdUsuarioAuditoria
        If mo_ReglasFarmacia.FarmMovimientoSeleccionarPorId(mo_farmMovimiento1) Then
            cmbUnidosis.Visible = True
            lblUnidosis.Visible = True
            mo_Formulario.HabilitarDeshabilitar cmbUnidosis, False
            mo_cmbUnidosis.BoundText = Trim(Str(mo_farmMovimiento1.IdAlmacenDestino))
            SSTab1.Tab = 2
            mo_cmbFarmaciaOrigen.BoundText = mo_DoFarmMovimiento.IdAlmacenOrigen
            FraUnidosis.Enabled = False
        End If
   End If
   'Impresion del FUA
   If Val(mo_cmbTipoFinanciamiento.BoundText) = sghTipoFinanciamiento.sghSis Then
      btnImprimeFichaSIS.Visible = True
   End If
End Sub
Function ValidarDatosObligatorios() As Boolean
   ValidarDatosObligatorios = False
   ms_MensajeError = ""
   If optPreventa.Value Then
        If cmbAlmOrigen.Text = "" Then
            ms_MensajeError = ms_MensajeError + "Por favor elija el Almacén Origen" + Chr(13)
        ElseIf cmbTipoFinanciamiento.Text = "" Then
            ms_MensajeError = ms_MensajeError + "Por favor elija PRODUCTO/PLAN" + Chr(13)
            cmbTipoFinanciamiento.SetFocus
        ElseIf cmbTipoReceta.Text = "" Then
            ms_MensajeError = ms_MensajeError + "Por favor elija el Tipo de Receta" + Chr(13)
            cmbTipoReceta.SetFocus
        End If
   ElseIf optVentas.Value Then
        If cmbAlmOrigen.Text = "" Then
            ms_MensajeError = ms_MensajeError + "Por favor elija el Almacén Origen" + Chr(13)
        ElseIf txtDatosDeCuenta.Text = "" Then
            ms_MensajeError = ms_MensajeError + "Por favor ingrese el N° de Cuenta" + Chr(13)
            txtNcuenta.SetFocus
        ElseIf txtNombrePaciente.Text = "" Then
            ms_MensajeError = ms_MensajeError + "Ese N° de Cuenta no tiene Paciente" + Chr(13)
            txtNhistoria.SetFocus
        ElseIf cmbTipoFinanciamiento.Text = "" Then
            ms_MensajeError = ms_MensajeError + "Por favor elija la PRODUCTO/PLAN" + Chr(13)
            cmbTipoFinanciamiento.SetFocus
        ElseIf cmbTipoReceta.Text = "" Then
            ms_MensajeError = ms_MensajeError + "Por favor elija el Tipo de Receta" + Chr(13)
            cmbTipoReceta.SetFocus
        ElseIf cmbPrescriptor.Text = "" Then
            ms_MensajeError = ms_MensajeError + "Por favor elija al PRESCRIPTOR" + Chr(13)
            cmbPrescriptor.SetFocus
        End If
   Else
        If cmbAlmOrigen.Text = "" Then
            ms_MensajeError = ms_MensajeError + "Por favor elija el Almacén Origen" + Chr(13)
        ElseIf cmbTipoFinanciamiento.Text = "" Then
            ms_MensajeError = ms_MensajeError + "Por favor elija la PRODUCTO/PLAN" + Chr(13)
            cmbTipoFinanciamiento.SetFocus
        ElseIf cmbTipoReceta.Text = "" Then
            ms_MensajeError = ms_MensajeError + "Por favor elija el Tipo de Receta" + Chr(13)
            cmbTipoReceta.SetFocus
        End If
   End If
   lnTotalDocumento = grdProductos.DevuelveTotal
   If optPreventa.Value = True Then
        If lnTotalDocumento <= 0 Then
           ms_MensajeError = ms_MensajeError + "El Total es MENOR o igual a CERO" + Chr(13)
        End If
   End If
   Set mRs_Productos = grdProductos.DevuelveProductos
   If mRs_Productos.RecordCount = 0 Then
       ms_MensajeError = ms_MensajeError + "Por favor Ingrese Productos" + Chr(13)
   Else
        mRs_Productos.MoveFirst
        Do While Not mRs_Productos.EOF
           If Trim(mRs_Productos.Fields!codigo) = "" Or Trim(mRs_Productos.Fields!nombreProducto) = "" Then
              mRs_Productos.Delete
              mRs_Productos.Update
           ElseIf IsNumeric(Trim(mRs_Productos.Fields!codigo)) Then 'Frank 04082015
                If Val(mRs_Productos.Fields!codigo) <= 0 Then
                    mRs_Productos.Delete
                    mRs_Productos.Update
                End If
           Else
              
           End If
            mRs_Productos.MoveNext
        Loop
        If mRs_Productos.RecordCount = 0 Then
            ms_MensajeError = ms_MensajeError + "Por favor Ingrese Productos" + Chr(13)
        Else
             mRs_Productos.MoveFirst
             Do While Not mRs_Productos.EOF
                If mRs_Productos.Fields!Cantidad <= 0 Or mRs_Productos!Cantidad > mRs_Productos!saldo Then
                   ms_MensajeError = ms_MensajeError + "El producto " + Trim(mRs_Productos.Fields!codigo) + " - " + Trim(mRs_Productos.Fields!nombreProducto) + "  Tiene problemas de Saldo" + Chr(13)
                End If
                mRs_Productos.Update
                mRs_Productos.MoveNext
             Loop
        End If
        'Es un despacho hacia la FARMACIA UNIDOSIS
        If ms_MensajeError = "" And cmbUnidosis.Visible = True Then
             If cmbUnidosis.Text = "" Then
                ms_MensajeError = ms_MensajeError + "Debe elegir la F.UNIDOSIS (destino)" + Chr(13)
             Else
                Dim lcCodigoConPunto As String
                Dim rs As New Recordset
                Dim oConexion As New Connection
                oConexion.CommandTimeout = 900
                oConexion.CursorLocation = adUseClient
                oConexion.Open SIGHEntidades.CadenaConexion
                If oRsItemsUnidosis.RecordCount = 0 Then
                   ms_MensajeError = ms_MensajeError + "No hay ITEMS en la FARMACIA UNIDOSIS" + Chr(13)
                Else
                   mRs_Productos.MoveFirst
                   Do While Not mRs_Productos.EOF
                      oRsItemsUnidosis.MoveFirst
                      oRsItemsUnidosis.Find "codigo='" & mRs_Productos!codigo & "'"
                      If oRsItemsUnidosis.EOF Then
                         ms_MensajeError = ms_MensajeError + "El ITEM " + Trim(mRs_Productos.Fields!codigo) + " - " + Trim(mRs_Productos.Fields!nombreProducto) + "  no pertenece a FARMACIA UNIDOSIS" + Chr(13)
                      Else
                         lcCodigoConPunto = Trim(mRs_Productos!codigo) & SIGHEntidades.Pto
                         Set rs = mo_ReglasFacturacion.FacturacionBienesPorCodigo(lcCodigoConPunto, _
                                                   Val(mo_cmbTipoFinanciamiento.BoundText), oConexion)
                         If rs.RecordCount = 0 Then
                            ms_MensajeError = ms_MensajeError + "El ITEM " + lcCodigoConPunto + " - " + Trim(oRsItemsUnidosis!Descripcion) + "  no tiene PRECIO" + Chr(13)
                         End If
                         rs.Close
                      End If
                      mRs_Productos.MoveNext
                   Loop
                End If
                oConexion.Close
                Set oConexion = Nothing
                Set rs = Nothing
             End If
        End If
        '
   End If
   If txtNombrePaciente.Text = "" And (mo_cmbTipoFinanciamiento.BoundText = "6" Or mo_cmbTipoFinanciamiento.BoundText = "10") Then
      'defensa nacional, credito personal
      ms_MensajeError = ms_MensajeError + "Debe elegir el PACIENTE para ese PRODUCTO/PLAN" + Chr(13)
   End If
   '
   Dim lbContinuarUnidosis As Boolean
   lbContinuarUnidosis = False
   If Me.optVentas.Value = True And cmbUnidosis.Visible = False And lbLaFarmaciaEsUnidosis = False Then    'solo RC6 si la FARMACIA origen NO ES UNIDOSIS
      lbContinuarUnidosis = True
   ElseIf Me.optVentas.Value = True And lbLaFarmaciaEsUnidosis = True Then
      lbContinuarUnidosis = False
   End If
   If lbContinuarUnidosis Then
      If mo_ReglasSISgalenhos.ReglasDeConsistenciaSISestanOK(sghVentasFarmacia, wxParametro302, _
                                                           ml_idFuenteFinanciamiento, ms_MensajeError, _
                                                           Val(Me.txtNcuenta.Text), "", "", lnIdTipoServicio, _
                                                           mi_Opcion, True, False, 0, Me.txtFregistro.Text, "", mRs_Productos) = True Then
      End If
      ms_MensajeError = ms_MensajeError & mo_ReglasSISgalenhos.ReglasDeConsistenciaSISsoloFarmaciaXmonto(Val(Me.txtNcuenta.Text), _
                                                  ml_idFuenteFinanciamiento, Me.txtFregistro.Text, mi_Opcion, mRs_Productos, False)

   End If
   '
   If ms_MensajeError <> "" Then
       MsgBox ms_MensajeError, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function


Sub CargaDatosAlObjetosDeDatos()
    Select Case mi_Opcion
    Case sghAgregar
        If optPreventa.Value Then
            With mo_DofarmPreVenta
                .fechaCreacion = txtFregistro.Text
                .HoraCreacion = lcBuscaParametro.RetornaHoraServidorSQL
                .IdAlmacen = Val(mo_cmbAlmacenOrigen.BoundText)
                .idCuentaAtencion = IIf(Len(Trim(txtDatosDeCuenta.Text)) > 0, Val(txtNcuenta.Text), 0)
                .idDiagnostico = ml_IdDiagnostico
                .idEstadoPreventa = sghEstadoTabla.sghRegistrado    'Por cancelar en Caja
                .IdPaciente = ml_IdPaciente
                .idPrescriptor = Val(mo_cmbPrescriptor.BoundText)
                .IdTipoFinanciamiento = Val(mo_cmbTipoFinanciamiento.BoundText)
                .idTipoReceta = Val(mo_cmbTipoReceta.BoundText)
                .idUsuario = ml_idUsuario
                .IdUsuarioAuditoria = ml_idUsuario
                .idVendedor = ml_idUsuario
                .Total = lnTotalDocumento
                If Val(txtRedondeo.Text) > 0 And CCur(txtRedondeo.Text) <> lnTotalDocumento Then
                      .Total = CCur(txtRedondeo.Text)
                End If
                .FechaHoraPrescribe = txtFprescribe.Text
                .dni = Me.txtNhistoria.Text
                .paciente = UCase(Me.txtNombrePaciente.Text)
            End With
        Else
            With mo_DoFarmMovimiento
                .fechaCreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL
                .IdAlmacenDestino = 0   '<<ninguno>>
                .IdAlmacenOrigen = Val(mo_cmbAlmacenOrigen.BoundText)
                .idEstadoMovimiento = sghEstadoTabla.sghRegistrado    'registrado
                .idTipoConcepto = ml_idTipoConcepto
                .idUsuario = ml_idUsuario
                .IdUsuarioAuditoria = ml_idUsuario
                .MovTipo = lcConstanteMovimientoSalida
                .Observaciones = txtObservaciones.Text
                .Total = lnTotalDocumento
                If Format(.fechaCreacion, "hh:mm:ss") = "00:00:00" Then
                   .fechaCreacion = CDate(Format(.fechaCreacion, "dd/mm/yyyy") & " 00:00:01")
                End If
            End With
            With mo_DoFarmMovimientoVentas
                .idCuentaAtencion = Val(txtNcuenta.Text)
                .idDiagnostico = ml_IdDiagnostico
                .IdPaciente = ml_IdPaciente
                .idPrescriptor = Val(mo_cmbPrescriptor.BoundText)
                .IdTipoFinanciamiento = Val(mo_cmbTipoFinanciamiento.BoundText)
                .idTipoReceta = Val(mo_cmbTipoReceta.BoundText)
                .IdUsuarioAuditoria = ml_idUsuario
                .MovTipo = lcConstanteMovimientoSalida
                .tipoVenta = lcConstanteVentaDirecta
                .idFuenteFinanciamiento = ml_idFuenteFinanciamiento
                .FechaHoraPrescribe = txtFprescribe.Text
            End With
        End If
   Case sghModificar
        If optPreventa.Value Then
            With mo_DofarmPreVenta
                .FechaModificacion = lcBuscaParametro.RetornaFechaServidorSQL
                .idDiagnostico = ml_IdDiagnostico
                .idPrescriptor = Val(mo_cmbPrescriptor.BoundText)
                .idTipoReceta = Val(mo_cmbTipoReceta.BoundText)
                .idUsuarioModifica = ml_idUsuario
                .Total = lnTotalDocumento
                .IdUsuarioAuditoria = ml_idUsuario
                If Val(txtRedondeo.Text) > 0 And CCur(txtRedondeo.Text) <> lnTotalDocumento Then
                      .Total = CCur(txtRedondeo.Text)
                End If
                .FechaHoraPrescribe = txtFprescribe.Text
                .dni = Me.txtNhistoria.Text
                .paciente = UCase(Me.txtNombrePaciente.Text)
                .IdPaciente = ml_IdPaciente
            End With
        Else
            With mo_DoFarmMovimiento
                .Observaciones = txtObservaciones.Text
                .Total = lnTotalDocumento
                .IdUsuarioAuditoria = ml_idUsuario
               ' .FechaCreacion = txtFregistro.Text
            End With
            With mo_DoFarmMovimientoVentas
                .idDiagnostico = ml_IdDiagnostico
                .idPrescriptor = Val(mo_cmbPrescriptor.BoundText)
                .idTipoReceta = Val(mo_cmbTipoReceta.BoundText)
                .IdUsuarioAuditoria = ml_idUsuario
                .idFuenteFinanciamiento = ml_idFuenteFinanciamiento
                .idTipoReceta = Val(mo_cmbTipoReceta.BoundText)
                .FechaHoraPrescribe = txtFprescribe.Text
            End With
        End If
   Case sghEliminar
        If optPreventa.Value Then
            With mo_DofarmPreVenta
                .idEstadoPreventa = sghEstadoTabla.sghAnulado    'Anulado
                .IdUsuarioAuditoria = ml_idUsuario
            End With
        Else
            With mo_DoFarmMovimiento
                .fechaAnulacion = lcBuscaParametro.RetornaFechaHoraServidorSQL
                .idEstadoMovimiento = sghEstadoTabla.sghAnulado   'Anulado
                .IdUsuarioAuditoria = ml_idUsuario
            End With
        End If
   End Select
End Sub

Function DevuelveItemsDeFarmaciaUNIDOSIS() As Recordset
    Dim mRs_Productos1 As New Recordset
    Dim rs As New Recordset
    Dim oConexion As New ADODB.Connection
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHEntidades.CadenaConexion
    With mRs_Productos1
          .Fields.Append "IdProducto", adInteger, 4
          .Fields.Append "Codigo", adVarChar, 20
          .Fields.Append "NombreProducto", adChar, 300
          .Fields.Append "idTipoSalidaBienInsumo", adInteger
          .Fields.Append "Lote", adVarChar, 15
          .Fields.Append "FechaVencimiento", adDate, , adFldIsNullable
          .Fields.Append "Cantidad", adInteger
          .Fields.Append "Precio", adDouble
          .Fields.Append "Total", adDouble
          .Fields.Append "MovNumeroS", adChar, 9, adFldIsNullable
          .Fields.Append "RegistroSanitario", adVarChar, 50, adFldIsNullable
          .Fields.Append "NumeroDocumento", adVarChar, 20, adFldIsNullable 'Frank 07082015
          .Fields.Append "esPaquete", adBoolean
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
    Set rs = mo_ReglasFarmacia.FarmMovimientosDetalleDevuelveTodosItems(oConexion, lcConstanteMovimientoEntrada, mo_farmMovimiento1.movNumero)
    Do While Not rs.EOF
        mRs_Productos1.AddNew
        mRs_Productos1!idProducto = rs!idProducto
        mRs_Productos1!codigo = rs!codigo
        mRs_Productos1!nombreProducto = rs!Nombre
        mRs_Productos1!Lote = rs!Lote
        mRs_Productos1!FechaVencimiento = rs!FechaVencimiento
        mRs_Productos1!Cantidad = rs!Cantidad
        mRs_Productos1!Precio = rs!Precio
        mRs_Productos1!Total = rs!Total
        mRs_Productos1!idTipoSalidaBienInsumo = rs!idTipoSalidaBienInsumo
        mRs_Productos1!registroSanitario = rs!registroSanitario
        If Not IsNull(rs!esPaquete) Then
          mRs_Productos1!esPaquete = rs!esPaquete
        End If
        mRs_Productos1.Update
        rs.MoveNext
    Loop
    Set DevuelveItemsDeFarmaciaUNIDOSIS = mRs_Productos1
    oConexion.Close
    Set oConexion = Nothing
    Set rs = Nothing
End Function

'*********Es un despacho hacia la FARMACIA UNIDOSIS*********
Sub CreaNIaFarmaciaUNIDOSIS(oRsProductosConLotes1 As Recordset)
    If lnCuentaUnidosis > 0 And cmbUnidosis.Visible = True Then
        Dim mo_farmMovimientoNotaIngreso1 As New DOfarmMovimientoNotaIngreso
        Dim oDoProveedores1 As New DoProveedores
        Dim mo_farmMovimiento2 As New farmMovimiento
        Dim oConexion As New Connection
        Dim rs As New Recordset
        Dim mo_FarmMovimientoNotaIngreso2 As New FarmMovimientoNotaIngreso
        Dim lnTotalUnidosis As Double, lnImporte As Double, lcCodigoConPunto As String
        Dim lnConvertir As Long, ActualizarDatos1 As Boolean
        oConexion.CommandTimeout = 900
        oConexion.CursorLocation = adUseClient
        oConexion.Open SIGHEntidades.CadenaConexion
        lnTotalUnidosis = 0
        If mi_Opcion <> sghEliminar Then
            oRsProductosConLotes1.MoveFirst
            Do While Not oRsProductosConLotes1.EOF
                oRsItemsUnidosis.MoveFirst
                oRsItemsUnidosis.Find "codigo='" & Trim(oRsProductosConLotes1!codigo) & "'"
                If Not oRsItemsUnidosis.EOF Then
                    lcCodigoConPunto = Trim(oRsProductosConLotes1!codigo) & SIGHEntidades.Pto
                    Set rs = mo_ReglasFacturacion.FacturacionBienesPorCodigo(lcCodigoConPunto, _
                                              Val(mo_cmbTipoFinanciamiento.BoundText), oConexion)
                    If rs.RecordCount > 0 Then
                        lnConvertir = Val(oRsItemsUnidosis!convertir)
                        lnImporte = Round(rs!PrecioUnitario * oRsProductosConLotes1!Cantidad * lnConvertir, 2)
                        oRsProductosConLotes1!idProducto = rs!idProducto
                        oRsProductosConLotes1!codigo = lcCodigoConPunto
                        oRsProductosConLotes1!Cantidad = oRsProductosConLotes1!Cantidad * lnConvertir
                        oRsProductosConLotes1!Precio = rs!PrecioUnitario
                        oRsProductosConLotes1!Total = lnImporte
                        oRsProductosConLotes1.Update
                        lnTotalUnidosis = lnTotalUnidosis + lnImporte
                    End If
                    rs.Close
                End If
                oRsProductosConLotes1.MoveNext
            Loop
        End If
        Select Case mi_Opcion
        Case sghEliminar, sghModificar
            Set mo_farmMovimiento2.Conexion = oConexion
            Set mo_FarmMovimientoNotaIngreso2.Conexion = oConexion
            mo_farmMovimiento1.IdUsuarioAuditoria = mo_DoFarmMovimiento.IdUsuarioAuditoria
            If mo_farmMovimiento1.movNumero <> "" Then
               mo_farmMovimientoNotaIngreso1.MovTipo = lcConstanteMovimientoEntrada
               mo_farmMovimientoNotaIngreso1.movNumero = mo_farmMovimiento1.movNumero
               mo_farmMovimientoNotaIngreso1.IdUsuarioAuditoria = mo_farmMovimiento1.IdUsuarioAuditoria
               If mo_FarmMovimientoNotaIngreso2.SeleccionarPorId(mo_farmMovimientoNotaIngreso1) Then
                    If mi_Opcion = sghModificar Then
                        mo_farmMovimiento1.Total = lnTotalUnidosis
                        ActualizarDatos1 = mo_ReglasFarmacia.ModificaDatosDeNotaIngreso(mo_farmMovimiento1, _
                                      mo_farmMovimientoNotaIngreso1, oDoProveedores1, oRsProductosConLotes1, _
                                      mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
                        If ActualizarDatos1 = True Then
                           MsgBox "Se actualizó Nota de Ingreso en FARMACIA UNIDOSIS en forma automática", vbInformation, Me.Caption
                        End If
                    Else
                        mo_farmMovimiento1.idEstadoMovimiento = sghEstadoTabla.sghAnulado
                        mo_farmMovimiento1.idUsuarioAnulacion = SIGHEntidades.Usuario
                        mo_farmMovimiento1.fechaAnulacion = lcBuscaParametro.RetornaFechaHoraServidorSQL
                        ActualizarDatos1 = mo_ReglasFarmacia.AnulaNotaIngreso(mo_farmMovimiento1, _
                                   mo_farmMovimientoNotaIngreso1, 0, oRsProductosConLotes1, mo_lnIdTablaLISTBARITEMS, _
                                   mo_lcNombrePc)
                        If ActualizarDatos1 = True Then
                           MsgBox "Se Anuló Nota de Ingreso en FARMACIA UNIDOSIS en forma automática", vbInformation, Me.Caption
                        End If
                    End If
               End If
            End If
        Case sghAgregar
            With mo_farmMovimiento1
                '.movNumero
                .MovTipo = lcConstanteMovimientoEntrada
                .IdAlmacenDestino = Val(mo_cmbUnidosis.BoundText)
                .idTipoConcepto = 4     '20
                .DocumentoIdtipo = 3    '10
                .DocumentoNumero = Format(Now, SIGHEntidades.DevuelveFechaSoloFormato_DMYHMS)
                .Total = lnTotalUnidosis
                .fechaCreacion = mo_DoFarmMovimiento.fechaCreacion
                .IdUsuarioAuditoria = mo_DoFarmMovimiento.IdUsuarioAuditoria
                .idUsuario = mo_DoFarmMovimiento.IdUsuarioAuditoria
                .idEstadoMovimiento = sghEstadoTabla.sghRegistrado
                
                .IdAlmacenOrigen = Val(mo_cmbFarmaciaOrigen.BoundText)
            End With
            With mo_farmMovimientoNotaIngreso1
                '.movNumero
                .MovTipo = lcConstanteMovimientoEntrada
                .DocumentoFechaRecepcion = mo_DoFarmMovimiento.fechaCreacion
                .OrigenIdTipo = 22
                .idTipoCompra = 1
                .idTipoProceso = 1
            End With
            With oDoProveedores1
            End With
            ActualizarDatos1 = mo_ReglasFarmacia.AgregaDatosDeNotaIngreso(mo_farmMovimiento1, _
                               mo_farmMovimientoNotaIngreso1, oDoProveedores1, oRsProductosConLotes1, 0, _
                               mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
            mo_DoFarmMovimiento.Observaciones = mo_farmMovimiento1.movNumero
            Set mo_farmMovimiento2.Conexion = oConexion
            If mo_farmMovimiento2.Modificar(mo_DoFarmMovimiento) Then
               MsgBox "Se creó Nota de Ingreso en forma automática en " & cmbUnidosis.Text, vbInformation, Me.Caption
            End If
        End Select
        oConexion.Close
        Set mo_farmMovimientoNotaIngreso1 = Nothing
        Set oDoProveedores1 = Nothing
        Set mo_farmMovimiento2 = Nothing
        Set oConexion = Nothing
        Set rs = Nothing
    End If
End Sub

Function AgregarDatos() As Boolean
    If optPreventa.Value Then
        AgregarDatos = mo_ReglasFarmacia.AgregaDatosDePreVenta(mo_DofarmPreVenta, mRs_Productos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lnIdReceta)
        'debb-16/02/2011
        txtNpreventa.Text = Trim(Str(mo_DofarmPreVenta.idPreVenta)) + lcEFE
    Else
        AgregarDatos = mo_ReglasFarmacia.AgregaDatosDeVentaDirecta(mo_DoFarmMovimiento, mo_DoFarmMovimientoVentas, mRs_Productos, _
                                         mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lnIdReceta, lnEpsPorcentaje)
        txtDocumento.Text = mo_DoFarmMovimiento.DocumentoNumero
        
        If mo_DoFarmMovimientoVentas.idCuentaAtencion > 0 Then
           If Val(wxParametro208) <> 7686 Or Val(wxParametro280) <> 2378 Then     'cajamarca/sicuani
                mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar mo_DoFarmMovimientoVentas.idCuentaAtencion, False, 0
                mo_ReglasSISgalenhos.FuaActualizaDespachosEnFarmacia mo_DoFarmMovimientoVentas.idCuentaAtencion, wxParametro302, lnIdTipoServicio, ml_idFuenteFinanciamiento
           End If
        End If
        If mo_ReglasFarmacia.IdOrdenPago > 0 Then
           lblOrdenPago.Caption = "N° Orden de Pago: " & mo_ReglasFarmacia.IdOrdenPago
        End If
        '
        CreaNIaFarmaciaUNIDOSIS mo_ReglasFarmacia.DevuelveProductosConLotes
    End If
    ms_MensajeError = mo_ReglasFarmacia.MensajeError
End Function
Function ModificarDatos() As Boolean
    If optPreventa.Value Then
        ModificarDatos = mo_ReglasFarmacia.ModificaDatosDePreVenta(mo_DofarmPreVenta, mRs_Productos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
        'debb-16/02/2011
        txtNpreventa.Text = Trim(txtNpreventa.Text) + lcEFE
    Else
        ModificarDatos = mo_ReglasFarmacia.ModificaDatosVentaDirecta(mo_DoFarmMovimiento, mo_DoFarmMovimientoVentas, _
                        mRs_Productos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lnIdReceta, lnEpsPorcentaje, Val(lblOrdenPago.Tag))
        If mo_DoFarmMovimientoVentas.idCuentaAtencion > 0 Then
           If Val(wxParametro208) <> 7686 Or Val(wxParametro280) <> 2378 Then  'cajamarca/sicuani
                mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar mo_DoFarmMovimientoVentas.idCuentaAtencion, False, 0
                mo_ReglasSISgalenhos.FuaActualizaDespachosEnFarmacia mo_DoFarmMovimientoVentas.idCuentaAtencion, wxParametro302, lnIdTipoServicio, ml_idFuenteFinanciamiento
           End If
        End If
        If mo_ReglasFarmacia.IdOrdenPago > 0 Then
           lblOrdenPago.Caption = "N° Orden de Pago: " & mo_ReglasFarmacia.IdOrdenPago
        End If
        '
        CreaNIaFarmaciaUNIDOSIS mo_ReglasFarmacia.DevuelveProductosConLotes
    End If
    ms_MensajeError = mo_ReglasFarmacia.MensajeError
End Function
Function Anular() As Boolean
    If optPreventa.Value Then
        Anular = mo_ReglasFarmacia.AnulaPreVenta(mo_DofarmPreVenta, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lnIdReceta)
        'debb-16/02/2011
        txtNpreventa.Text = Trim(txtNpreventa.Text) + lcEFE
    Else
        Anular = mo_ReglasFarmacia.AnulaNotaSalida(mo_DoFarmMovimiento, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lnIdReceta, Val(lblOrdenPago.Tag), mo_DoFarmMovimientoVentas)
        If mo_DoFarmMovimientoVentas.idCuentaAtencion > 0 Then
           mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar mo_DoFarmMovimientoVentas.idCuentaAtencion, False, 0
           mo_ReglasSISgalenhos.FuaActualizaDespachosEnFarmacia mo_DoFarmMovimientoVentas.idCuentaAtencion, wxParametro302, lnIdTipoServicio, ml_idFuenteFinanciamiento
        End If
        '
        CreaNIaFarmaciaUNIDOSIS DevuelveItemsDeFarmaciaUNIDOSIS
    End If
    ms_MensajeError = mo_ReglasFarmacia.MensajeError
End Function
Private Sub btnCancelar_Click()
   If SIGHEntidades.ParaAuditoria = "" Then
      Me.Visible = False
      LimpiarVariablesDeMemoria
   ElseIf MsgBox("Hubo cambios, desea salir de todas maneras ?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
      Me.Visible = False
      LimpiarVariablesDeMemoria
      SIGHEntidades.ParaAuditoriaPorCadaDato sghAudLimpiar, ""
   End If
End Sub




Private Sub txtNhistoria_LostFocus()
      If mo_Teclado.TextoEsSoloNumeros(txtNhistoria.Text) And txtDatosDeCuenta.Text = "" Then
        Dim oRsTmp1 As New ADODB.Recordset
        Dim oDOPaciente As New sighComun.DOPaciente
        If Me.optPreventa.Value = True Then
           oDOPaciente.IdDocIdentidad = 1
           oDOPaciente.nroDocumento = txtNhistoria.Text
        Else
           oDOPaciente.NroHistoriaClinica = HCigualDNI_AgregaNUEVEaLaHistoria(txtNhistoria.Text)
        End If
        Set oRsTmp1 = mo_AdminAdmision.PacientesFiltrar(oDOPaciente, False, False, "")
        If oRsTmp1.RecordCount > 0 Then
           ml_IdPaciente = oRsTmp1.Fields!IdPaciente
           txtNombrePaciente.Text = Trim(oRsTmp1.Fields!ApellidoPaterno) & " " & Trim(oRsTmp1.Fields!ApellidoMaterno) & " " & oRsTmp1.Fields!PrimerNombre
        Else
           oRsTmp1.Close
           Set oRsTmp1 = mo_ReglasFarmacia.FarmPreventaFiltrar("dni='" & txtNhistoria.Text & "'")
           If oRsTmp1.RecordCount > 0 Then
               ml_IdPaciente = 0
               txtNombrePaciente.Text = oRsTmp1!paciente
           Else
                ml_IdPaciente = 0
                txtNombrePaciente.Text = ""
           End If
        End If
        oRsTmp1.Close
        Set oRsTmp1 = Nothing
        Set oDOPaciente = Nothing
      End If

End Sub

Private Sub txtObservaciones_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtObservaciones

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
        Case vbKeyEscape
'           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       Case vbKeyF3
           btnNuevo_Click
       End Select
End Sub

Sub LimpiarDatos()
   SIGHEntidades.ParaAuditoriaPorCadaDato sghAudLimpiar, ""
   lbCuentaDeEmergenciaCerrada = False
   cmbUnidosis.Visible = False
   lblUnidosis.Visible = False
   lblOrdenPago.Caption = ""
   lnEpsPorcentaje = 0
   txtNcuenta.Text = ""
   txtDatosDeCuenta.Text = ""
   txtPlan.Text = ""
   txtNhistoria.Text = ""
   txtNombrePaciente.Text = ""
   ml_IdPaciente = 0
   ml_idFuenteFinanciamiento = 0
   cmbTipoFinanciamiento.Text = ""
   cmbTipoReceta.Text = ""
   txtTurno.Text = ""
   txtCaja.Text = ""
   txtCajero.Text = ""
   ml_IdCajero = 0
   txtVendedor.Text = ""
   ml_IdVendedor = 0
   cmbPrescriptor.Text = ""
   ml_IdDiagnostico = 0
   txtDx.Text = ""
   txtNombreDx.Text = ""
   txtObservaciones.Text = ""
   ml_movNumero = ""
   txtDocumento.Text = ""
   txtNpreventa.Text = ""
   lnTotalDocumento = 0
   grdProductos.movNumero = 0
   chkPlanNoCubre.Value = 0
   lnIdTipoServicio = 0
   txtNreceta.Text = ""
   txtFprescribe.Text = lcBuscaParametro.RetornaFechaHoraServidorSQL
   Set grdHistorico.DataSource = Nothing
   Me.Height = 8085
   grdHistorico.Visible = False
   txtHtotCantidad.Visible = False
   grdProductos.PermiteAgregarItems = True
   grdProductos.LimpiarGrilla
   grdProductos.AgregaRegistro
   If optVentas.Value = True Then
      optVentas_Click 1
      On Error Resume Next
      txtNcuenta.SetFocus
   Else
      optPreventa_Click 1
      On Error Resume Next
      cmbTipoReceta.SetFocus
   End If
   
   'Me.KeyPreview = True
End Sub





Private Sub txtRedondeo_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
           KeyAscii = 0
       End If
End Sub

Private Sub txtRedondeo_LostFocus()
     If CCur(txtRedondeo.Text) - grdProductos.DevuelveTotal > 0.2 Then
        MsgBox "El redondeo es mayor de 0.20", vbInformation, Me.Caption
        txtRedondeo.Text = grdProductos.DevuelveTotal
        Exit Sub
     End If
     If grdProductos.DevuelveTotal - CCur(txtRedondeo.Text) > 0.2 Then
        MsgBox "El redondeo es menor de 0.20 ", vbInformation, Me.Caption
        txtRedondeo.Text = grdProductos.DevuelveTotal
        Exit Sub
     End If
End Sub

Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_Teclado = Nothing
    Set mo_Formulario = Nothing
    Set mo_cmbAlmacenOrigen = Nothing
    Set mo_cmbPrescriptor = Nothing
    Set mo_cmbTipoFinanciamiento = Nothing
    Set mo_cmbTipoReceta = Nothing
    Set mo_ReglasFarmacia = Nothing
    Set mo_AdminServiciosComunes = Nothing
    Set mo_AdminAdmision = Nothing
    Set mo_ReglasCaja = Nothing
    Set mo_ReglasFacturacion = Nothing
    Set oRsTipoFinanciamiento = Nothing
    Set mo_DofarmPreVenta = Nothing
    Set mo_DoPaciente = Nothing
    Set lcBuscaParametro = Nothing
    Set mRs_Productos = Nothing
    Set mo_DoFarmMovimiento = Nothing
    Set mo_DoFarmMovimientoVentas = Nothing
End Sub


Private Sub txtNreceta_LostFocus()
    If Val(txtNreceta.Text) > 0 Then
       Dim lcSql As String
       Dim oRsTmp1 As New Recordset, lnRecetaProcesada As Long, lnCuenta As Long, lcDespacho As String
       Dim oConexion As New Connection
       
       lnRecetaProcesada = Val(txtNreceta.Text)
       '
       oConexion.CommandTimeout = 300
       oConexion.CursorLocation = adUseClient
       oConexion.Open SIGHEntidades.CadenaConexion
       '
       Set oRsTmp1 = mo_ReglasComunes.RecetaCabeceraDetalleSeleccionaPorNroReceta(Val(txtNreceta.Text))
       
       If oRsTmp1.RecordCount > 0 Then
            If oRsTmp1.Fields!idEstado <> sghRecetaEstados.sighRecetaRegistrada Then
                lcDespacho = IIf(IsNull(oRsTmp1.Fields!DocumentoDespacho), "", oRsTmp1.Fields!DocumentoDespacho)
                mo_ReglasComunes.RecetaChequeaEstadoActual oRsTmp1.Fields!idCuentaAtencion, _
                                                           oRsTmp1.Fields!idEstado, _
                                                           0, lcDespacho
                txtNreceta.Text = ""
            Else
                
                If Not IsNull(oRsTmp1.Fields!idMedicoReceta) Then
                   mo_cmbPrescriptor.BoundText = oRsTmp1.Fields!idMedicoReceta
                End If
                'debb-24/06/2015
                lcSql = ""
                If Not IsNull(oRsTmp1!fechaVigencia) Then
                   If CDate(lcBuscaParametro.RetornaFechaServidorSQL) > oRsTmp1!fechaVigencia Then
                      lcSql = "Esa Receta tiene FECHA DE VIGENCIA: " & oRsTmp1!fechaVigencia
                      MsgBox lcSql, vbInformation, Me.Caption
                      txtNreceta.Text = ""
                   End If
                End If
                If lcSql = "" Then     'debb-24/06/2015
                    lbCuentaDeEmergenciaCerrada = mo_ReglasComunes.CuentaDeEmergenciaCerrada(oRsTmp1!idCuentaAtencion, sghPtoCargaFarmacia)
                    txtFprescribe.Text = Format(oRsTmp1.Fields!fechaReceta, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HM)
                    If mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(oRsTmp1.Fields!IdFormaPago, oConexion) = True And oRsTmp1.Fields!IdTipoServicio = 1 Then
                         optPreventa.Value = True
                         optPreventa_Click True
                         txtNcuenta.Text = oRsTmp1.Fields!idCuentaAtencion
                         txtNcuenta_LostFocus
                         grdProductos.CargaProductosPorIdReceta oRsTmp1
                         grdProductos.esReceta = True
                         lnIdReceta = lnRecetaProcesada
                    Else
                         optVentas.Value = True
                         optVentas_Click True
                         txtNcuenta.Text = oRsTmp1.Fields!idCuentaAtencion
                         txtNcuenta_LostFocus
                         'If lbCuentaDeEmergenciaCerrada = False Then
                         grdProductos.CargaProductosPorIdReceta oRsTmp1
                         grdProductos.esReceta = True
                         lnIdReceta = lnRecetaProcesada
                         'End If
                    End If
                    grdProductos.PermiteAgregarItems = False
                End If
            End If
       Else
            MsgBox "Ese N° Receta NO EXISTE", vbInformation, "Caja"
            txtNreceta.Text = ""
       End If
       If oRsTmp1.State = 1 Then oRsTmp1.Close
       Set oRsTmp1 = Nothing
    End If
    On Error Resume Next
    oConexion.Close
    Set oConexion = Nothing
End Sub
Private Sub txtNreceta_KeyDown(KeyCode As Integer, Shift As Integer)
       mo_Teclado.RealizarNavegacion KeyCode, txtNreceta
       AdministrarKeyPreview KeyCode
End Sub

'debb-09/07/2015
Private Sub ImprimeDocumento()
    Dim oRptClase As New rCrystal
    Dim oDOfarmAlmacen As New DoFarmAlmacen
    Set oDOfarmAlmacen = mo_ReglasFarmacia.FarmAlmacenSeleccionarPorId(Val(mo_cmbAlmacenOrigen.BoundText))
    oRptClase.MovTipo = "S"
    oRptClase.Documento = mo_DoFarmMovimiento.movNumero
    oRptClase.TextoDelFiltro = "VENTAS"
    oRptClase.Almacen = "Paciente: " & HCigualDNI_DevuelveHistoriaConCerosIzquierda(txtNhistoria.Text, False) & " - " & Trim(txtNombrePaciente.Text) & _
                        "      (Dx : " & Trim(Me.txtDx.Text) & " - " & Trim(txtNombreDx.Text) & ")"
    oRptClase.AlmacenO = "(" & oDOfarmAlmacen.CodigoSismed & ")" & cmbAlmOrigen.Text
    oRptClase.HoraInicio = txtFregistro.Text & " " & Me.txtHoraRegistro.Text
    oRptClase.HoraFin = Trim(txtTipoComprobante.Text) & ": " & Trim(txtDocumento.Text)
    oRptClase.Importe = CCur(txtRedondeo.Text)
    oRptClase.TipoReporte = "NiNs"
    oRptClase.Observaciones = Trim(Left(txtPlan.Text, 29)) & _
                              "      NªCuenta: " & Trim(txtNcuenta.Text) & " " & Trim(txtDatosDeCuenta.Text)
    
    oRptClase.EsUnaDonacion = False
    'If Trim(cmbTipoDocum.Text) <> "" Then
    '    oRptClase.Proveedor = Trim(cmbTipoDocum.Text) & "/" & Trim(txtNdocum.Text)
    'End If
    oRptClase.idUsuario = ml_idUsuario
    oRptClase.Show vbModal
    Set oRptClase = Nothing
    Set oDOfarmAlmacen = Nothing
End Sub

Sub CargaItemsDebajoDeStockMinimo()
 Dim oRsTmp As New Recordset
 CargaInventarioExcel.Enabled = True
 grdHistorico.Caption = "Lista de Medicamentos/Insumos por debajo de su STOCK MINIMO"
 Set oRsTmp = mo_ReglasFarmacia.FarmaciaItemsPorDebajoStockMinimo
 If cmbAlmOrigen.Text <> "" Then
    oRsTmp.Filter = "idAlmacen=" & mo_cmbAlmacenOrigen.BoundText
 End If
 Set grdHistorico.DataSource = oRsTmp
 grdHistorico.Visible = True
 Me.Height = 9930
End Sub
