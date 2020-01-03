VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.UserControl ucEstadoCuenta 
   ClientHeight    =   9855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12705
   ScaleHeight     =   9855
   ScaleWidth      =   12705
   Begin VB.CommandButton cmdLiquidacion 
      Caption         =   "Imprimir RESUMEN DE LIQUIDACION"
      Height          =   1005
      Left            =   7375
      Picture         =   "ucEstadoCuenta.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   164
      Top             =   8460
      Width           =   1695
   End
   Begin UltraGrid.SSUltraGrid grdCuentasPorTipoServicio 
      Height          =   225
      Left            =   60
      TabIndex        =   109
      Top             =   1740
      Visible         =   0   'False
      Width           =   12510
      _ExtentX        =   22066
      _ExtentY        =   397
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
      Caption         =   "Lista de Pacientes"
   End
   Begin VB.CommandButton btnAceptar 
      Caption         =   "Actualiza SIS, SOAT, Exon., Conven. [F2]"
      DisabledPicture =   "ucEstadoCuenta.ctx":04D9
      DownPicture     =   "ucEstadoCuenta.ctx":0939
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
      Height          =   1005
      Left            =   10935
      Picture         =   "ucEstadoCuenta.ctx":0DAE
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8460
      Width           =   1695
   End
   Begin TabDlg.SSTab ucFacturacionBienesInsumos 
      Height          =   5475
      Left            =   60
      TabIndex        =   7
      Top             =   2940
      Width           =   12570
      _ExtentX        =   22172
      _ExtentY        =   9657
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
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
      TabCaption(0)   =   "Servicios"
      TabPicture(0)   =   "ucEstadoCuenta.ctx":1223
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTotalServicios"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblTotalSeguroServicio"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label49"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ucFacturacionServicios"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtTotalServicios"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkServiciosTodos"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtTotalSeguroServicio"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtPagosAdelantoS"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Farmacia"
      TabPicture(1)   =   "ucEstadoCuenta.ctx":123F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtPagosAdelantoF"
      Tab(1).Control(1)=   "txtTotalSeguroFarmacia"
      Tab(1).Control(2)=   "chkFarmaciaTodos"
      Tab(1).Control(3)=   "txtTotalFarmacia"
      Tab(1).Control(4)=   "ucFacturacionBienes"
      Tab(1).Control(5)=   "Label50"
      Tab(1).Control(6)=   "lblTotalSeguroFarmacia"
      Tab(1).Control(7)=   "lblPagoFarmacia"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Consolidado"
      TabPicture(2)   =   "ucEstadoCuenta.ctx":125B
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtTotalApagar"
      Tab(2).Control(1)=   "txtPagosAdelantoC"
      Tab(2).Control(2)=   "txtTotalConsumo"
      Tab(2).Control(3)=   "txtExoneraciones"
      Tab(2).Control(4)=   "txtTotalSeguro"
      Tab(2).Control(5)=   "txtDevoluciones"
      Tab(2).Control(6)=   "grdCabecera"
      Tab(2).Control(7)=   "grdDetalle"
      Tab(2).Control(8)=   "Label27"
      Tab(2).Control(9)=   "Label51"
      Tab(2).Control(10)=   "Label52"
      Tab(2).Control(11)=   "Label53"
      Tab(2).Control(12)=   "Label54"
      Tab(2).Control(13)=   "Label60"
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "Resumen"
      TabPicture(3)   =   "ucEstadoCuenta.ctx":1277
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).Control(1)=   "Frame1"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Reembolso"
      TabPicture(4)   =   "ucEstadoCuenta.ctx":1293
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtPorReembolsar"
      Tab(4).Control(1)=   "txtReembolsoT"
      Tab(4).Control(2)=   "txtReembolsoS"
      Tab(4).Control(3)=   "txtReembolsoF"
      Tab(4).Control(4)=   "grdReembolsoF"
      Tab(4).Control(5)=   "Frame4"
      Tab(4).Control(6)=   "Frame2"
      Tab(4).Control(7)=   "Frame5"
      Tab(4).Control(8)=   "Frame3"
      Tab(4).Control(9)=   "Label55"
      Tab(4).Control(10)=   "Label57"
      Tab(4).Control(11)=   "lblTiempoDeCargaDeCuenta"
      Tab(4).ControlCount=   12
      TabCaption(5)   =   "Farmacia-Donaciones"
      TabPicture(5)   =   "ucEstadoCuenta.ctx":12AF
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "txtTotalDonaciones"
      Tab(5).Control(1)=   "grdItemsDonaciones"
      Tab(5).ControlCount=   2
      Begin VB.TextBox txtTotalDonaciones 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -64170
         TabIndex        =   210
         Top             =   4680
         Width           =   1185
      End
      Begin VB.TextBox txtPorReembolsar 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -66825
         TabIndex        =   205
         Top             =   4980
         Width           =   1185
      End
      Begin VB.TextBox txtReembolsoT 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -64200
         TabIndex        =   204
         Top             =   5010
         Width           =   1185
      End
      Begin VB.TextBox txtReembolsoS 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -64200
         TabIndex        =   203
         Top             =   4560
         Width           =   1185
      End
      Begin VB.TextBox txtReembolsoF 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -65430
         TabIndex        =   202
         Top             =   4560
         Width           =   1185
      End
      Begin VB.Frame Frame6 
         Caption         =   "RECALCULO de Cuenta de Atención"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4965
         Left            =   -66210
         TabIndex        =   193
         Top             =   360
         Width           =   3705
         Begin VB.CheckBox chkSoatParticular 
            Caption         =   "Pasa de Soat hacia PARTICULAR con Precios SOAT"
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
            Left            =   90
            TabIndex        =   196
            Top             =   4410
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.TextBox txtRecalculo 
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   1395
            Left            =   90
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            TabIndex        =   195
            Text            =   "ucEstadoCuenta.ctx":12CB
            Top             =   210
            Width           =   3525
         End
         Begin VB.CommandButton btnRecalculaPlan 
            Caption         =   "Cambia a otra 'Fuente Finanaciamiento/IAFA'"
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
            Height          =   525
            Left            =   90
            TabIndex        =   194
            Top             =   2490
            Visible         =   0   'False
            Width           =   3495
         End
         Begin MSDataListLib.DataCombo cmbFuenteFinanciamiento 
            Height          =   345
            Left            =   1800
            TabIndex        =   197
            Top             =   1680
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   609
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   "DataCombo1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo cmbFormaPago 
            Height          =   345
            Left            =   1800
            TabIndex        =   198
            Top             =   2040
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   609
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   "DataCombo1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label59 
            Caption         =   "Nuev.Producto/Plan"
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
            Left            =   90
            TabIndex        =   201
            Top             =   2100
            Width           =   1695
         End
         Begin VB.Label Label58 
            Caption         =   "Nuevo IAFA"
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
            Left            =   90
            TabIndex        =   200
            Top             =   1740
            Width           =   945
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            Caption         =   "T.Finan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -870
            TabIndex        =   199
            Top             =   2220
            Width           =   90
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "ALTA ADMINISTRATIVA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4965
         Left            =   -74880
         TabIndex        =   180
         Top             =   360
         Width           =   8685
         Begin VB.CommandButton btnCerrarCuenta 
            Caption         =   "Cerrar Cuenta"
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
            Height          =   525
            Left            =   60
            TabIndex        =   192
            Top             =   2565
            Width           =   1395
         End
         Begin VB.CommandButton btnAbrirCuenta 
            Caption         =   "Abrir Cuenta"
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
            Height          =   525
            Left            =   60
            TabIndex        =   191
            Top             =   1815
            Width           =   1395
         End
         Begin VB.CommandButton btnCtaPagada 
            Caption         =   "Cuenta Pagada"
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
            Height          =   525
            Left            =   60
            TabIndex        =   190
            Top             =   210
            Width           =   1395
         End
         Begin VB.CommandButton btnCtaAnulada 
            Caption         =   "Cuenta Anulada"
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
            Height          =   525
            Left            =   60
            TabIndex        =   189
            Top             =   3330
            Width           =   1395
         End
         Begin VB.TextBox txtCtaPagada 
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   615
            Left            =   1500
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            TabIndex        =   188
            Text            =   "ucEstadoCuenta.ctx":12D1
            Top             =   210
            Width           =   7125
         End
         Begin VB.TextBox txtCtaAbrir 
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   615
            Left            =   1500
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            TabIndex        =   187
            Text            =   "ucEstadoCuenta.ctx":12D7
            Top             =   1830
            Width           =   7125
         End
         Begin VB.TextBox txtCtaCerrar 
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   615
            Left            =   1500
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            TabIndex        =   186
            Text            =   "ucEstadoCuenta.ctx":12DD
            Top             =   2595
            Width           =   7125
         End
         Begin VB.TextBox txtCtaAnulada 
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   645
            Left            =   1500
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            TabIndex        =   185
            Text            =   "ucEstadoCuenta.ctx":12E3
            Top             =   3360
            Width           =   7125
         End
         Begin VB.TextBox txtPendienteSeguro 
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   735
            Left            =   1500
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            TabIndex        =   184
            Text            =   "ucEstadoCuenta.ctx":12E9
            Top             =   945
            Width           =   7125
         End
         Begin VB.CommandButton btnPendientePagoSeguro 
            Caption         =   "Pendiente Pago Seguros"
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
            Height          =   525
            Left            =   60
            TabIndex        =   183
            Top             =   930
            Width           =   1395
         End
         Begin VB.TextBox txtCtaConGarante 
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   645
            Left            =   1500
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            TabIndex        =   182
            Text            =   "ucEstadoCuenta.ctx":12F1
            Top             =   4200
            Width           =   7125
         End
         Begin VB.CommandButton btnCtaGarante 
            Caption         =   "Cuenta  con Garante"
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
            Height          =   525
            Left            =   60
            TabIndex        =   181
            Top             =   4170
            Width           =   1395
         End
      End
      Begin VB.TextBox txtTotalApagar 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -63840
         TabIndex        =   171
         Top             =   2580
         Width           =   945
      End
      Begin VB.TextBox txtPagosAdelantoC 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -71220
         TabIndex        =   170
         Text            =   "0"
         Top             =   2580
         Width           =   915
      End
      Begin VB.TextBox txtTotalConsumo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -73620
         TabIndex        =   169
         Text            =   "0"
         Top             =   2580
         Width           =   825
      End
      Begin VB.TextBox txtExoneraciones 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -69480
         TabIndex        =   168
         Text            =   "0"
         Top             =   2580
         Width           =   915
      End
      Begin VB.TextBox txtTotalSeguro 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -65880
         TabIndex        =   167
         Text            =   "0"
         Top             =   2580
         Width           =   1005
      End
      Begin VB.TextBox txtDevoluciones 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -67800
         TabIndex        =   166
         Text            =   "0"
         Top             =   2580
         Width           =   1005
      End
      Begin VB.TextBox txtPagosAdelantoF 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -69960
         TabIndex        =   115
         Text            =   "0"
         Top             =   4950
         Width           =   1005
      End
      Begin VB.TextBox txtPagosAdelantoS 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4980
         TabIndex        =   113
         Text            =   "0"
         Top             =   4950
         Width           =   1005
      End
      Begin VB.TextBox txtTotalSeguroFarmacia 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -66900
         TabIndex        =   22
         Text            =   "0"
         Top             =   4950
         Width           =   1005
      End
      Begin VB.TextBox txtTotalSeguroServicio 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8070
         TabIndex        =   20
         Text            =   "0"
         Top             =   4950
         Width           =   1005
      End
      Begin VB.CheckBox chkFarmaciaTodos 
         Caption         =   "Todos/Ninguno"
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
         Left            =   -74880
         TabIndex        =   19
         Top             =   4980
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.CheckBox chkServiciosTodos 
         Caption         =   "Todos/Ninguno"
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
         Left            =   150
         TabIndex        =   18
         Top             =   5040
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.TextBox txtTotalFarmacia 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -63540
         TabIndex        =   17
         Top             =   4980
         Width           =   945
      End
      Begin VB.TextBox txtTotalServicios 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   11430
         TabIndex        =   15
         Top             =   4950
         Width           =   1005
      End
      Begin SISGalenPlus.ucFactItemsEstadoCuenta ucFacturacionServicios 
         Height          =   4485
         Left            =   120
         TabIndex        =   129
         Top             =   390
         Width           =   12270
         _ExtentX        =   21643
         _ExtentY        =   8281
      End
      Begin SISGalenPlus.ucFactItemsEstadoCuenta ucFacturacionBienes 
         Height          =   4515
         Left            =   -74880
         TabIndex        =   130
         Top             =   390
         Width           =   12270
         _ExtentX        =   21643
         _ExtentY        =   8334
      End
      Begin UltraGrid.SSUltraGrid grdCabecera 
         Height          =   2085
         Left            =   -74850
         TabIndex        =   172
         Top             =   420
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   3678
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
         Caption         =   "Resumen por Punto de Carga"
      End
      Begin UltraGrid.SSUltraGrid grdDetalle 
         Height          =   2385
         Left            =   -74850
         TabIndex        =   173
         Top             =   3000
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   4207
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
         Caption         =   "...."
      End
      Begin UltraGrid.SSUltraGrid grdReembolsoF 
         Height          =   3945
         Left            =   -74880
         TabIndex        =   206
         Top             =   450
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   6959
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
         Caption         =   "Reembolsos"
      End
      Begin VB.Frame Frame4 
         Caption         =   "TOTAL SERVICIOS"
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
         Height          =   2085
         Left            =   -74790
         TabIndex        =   155
         Top             =   600
         Visible         =   0   'False
         Width           =   3420
         Begin VB.TextBox txtIngresadoServ 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1980
            TabIndex        =   159
            Top             =   270
            Width           =   1215
         End
         Begin VB.TextBox txtPendientePagoServ 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1980
            TabIndex        =   158
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtExoneradoServ 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1980
            TabIndex        =   157
            Top             =   1170
            Width           =   1215
         End
         Begin VB.TextBox TxtDctosServicio 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1980
            TabIndex        =   156
            Top             =   1590
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "INGRESADO (S/.)"
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
            Left            =   60
            TabIndex        =   163
            Top             =   360
            Width           =   1515
         End
         Begin VB.Label Label2 
            Caption         =   "PEND. PAGO (S/.)"
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
            Left            =   60
            TabIndex        =   162
            Top             =   810
            Width           =   1545
         End
         Begin VB.Label ll 
            Caption         =   "EXONERADO (S/.)"
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
            Left            =   60
            TabIndex        =   161
            Top             =   1230
            Width           =   1605
         End
         Begin VB.Label Label20 
            Caption         =   "DCTOS (S/.)"
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
            Left            =   60
            TabIndex        =   160
            Top             =   1650
            Width           =   1605
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "TOTAL BIENES E INSUMOS"
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
         Height          =   1725
         Left            =   -72480
         TabIndex        =   148
         Top             =   690
         Visible         =   0   'False
         Width           =   3405
         Begin VB.TextBox txtExoneradoBien 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1980
            TabIndex        =   151
            Top             =   1170
            Width           =   1215
         End
         Begin VB.TextBox txtPendientePagoBien 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1980
            TabIndex        =   150
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtIngresadoBien 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1980
            TabIndex        =   149
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label12 
            Caption         =   "EXONERADO (S/.)"
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
            Left            =   60
            TabIndex        =   154
            Top             =   1230
            Width           =   1605
         End
         Begin VB.Label Label13 
            Caption         =   "PEND. PAGO (S/.)"
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
            Left            =   60
            TabIndex        =   153
            Top             =   810
            Width           =   1545
         End
         Begin VB.Label Label14 
            Caption         =   "INGRESADO (S/.)"
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
            Left            =   60
            TabIndex        =   152
            Top             =   360
            Width           =   1515
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "TOTAL PAGADO"
         Enabled         =   0   'False
         Height          =   1845
         Left            =   -69870
         TabIndex        =   132
         Top             =   840
         Visible         =   0   'False
         Width           =   3405
         Begin VB.TextBox txtTotalPagado 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1965
            TabIndex        =   135
            Top             =   1260
            Width           =   1215
         End
         Begin VB.TextBox txtTotalBienPagado 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1980
            TabIndex        =   134
            Top             =   780
            Width           =   1215
         End
         Begin VB.TextBox txtTotalServPagado 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1980
            TabIndex        =   133
            Top             =   330
            Width           =   1215
         End
         Begin VB.Label Label15 
            Caption         =   "TOTAL"
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
            Left            =   195
            TabIndex        =   138
            Top             =   1290
            Width           =   1425
         End
         Begin VB.Label Label16 
            Caption         =   "TOTAL BIENES (S/.)"
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
            Left            =   180
            TabIndex        =   137
            Top             =   840
            Width           =   1425
         End
         Begin VB.Label Label17 
            Caption         =   "TOTAL SERVICIOS (S/.)"
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
            Left            =   180
            TabIndex        =   136
            Top             =   390
            Width           =   1785
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "TOTALES POR PAGAR"
         Enabled         =   0   'False
         Height          =   2265
         Left            =   -66420
         TabIndex        =   139
         Top             =   660
         Visible         =   0   'False
         Width           =   3405
         Begin VB.TextBox txtPagoACuentaServ 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1980
            TabIndex        =   143
            Top             =   1230
            Width           =   1215
         End
         Begin VB.TextBox txtTotalServ 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1980
            TabIndex        =   142
            Top             =   330
            Width           =   1215
         End
         Begin VB.TextBox txtTotalBien 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1980
            TabIndex        =   141
            Top             =   780
            Width           =   1215
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1980
            TabIndex        =   140
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "PAGO A CUENTA"
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
            Left            =   210
            TabIndex        =   147
            Top             =   1290
            Width           =   1425
         End
         Begin VB.Label Label10 
            Caption         =   "TOTAL SERVICIOS (S/.)"
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
            Left            =   180
            TabIndex        =   146
            Top             =   390
            Width           =   1785
         End
         Begin VB.Label Label9 
            Caption         =   "TOTAL BIENES (S/.)"
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
            Left            =   180
            TabIndex        =   145
            Top             =   840
            Width           =   1425
         End
         Begin VB.Label Label11 
            Caption         =   "TOTAL"
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
            Left            =   210
            TabIndex        =   144
            Top             =   1710
            Width           =   1425
         End
      End
      Begin UltraGrid.SSUltraGrid grdItemsDonaciones 
         Height          =   3915
         Left            =   -74820
         TabIndex        =   211
         Top             =   600
         Width           =   12060
         _ExtentX        =   21273
         _ExtentY        =   6906
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   71303188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   "ucEstadoCuenta.ctx":12F7
         Caption         =   "Productos"
      End
      Begin VB.Label Label55 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Por Reembolsar"
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
         Left            =   -68145
         TabIndex        =   209
         Top             =   5040
         Width           =   1260
      End
      Begin VB.Label Label57 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Reembolso Total"
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
         Left            =   -65625
         TabIndex        =   208
         Top             =   5070
         Width           =   1365
      End
      Begin VB.Label lblTiempoDeCargaDeCuenta 
         Caption         =   "......"
         Height          =   345
         Left            =   -74730
         TabIndex        =   207
         Top             =   4830
         Width           =   3765
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Saldo Final"
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
         Left            =   -64740
         TabIndex        =   179
         Top             =   2640
         Width           =   840
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Pagos a Cuenta"
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
         Left            =   -72555
         TabIndex        =   178
         Top             =   2640
         Width           =   1275
      End
      Begin VB.Label Label52 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Consumo"
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
         Left            =   -74880
         TabIndex        =   177
         Top             =   2640
         Width           =   1230
      End
      Begin VB.Label Label53 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Exoner"
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
         Left            =   -70110
         TabIndex        =   176
         Top             =   2640
         Width           =   570
      End
      Begin VB.Label Label54 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Seguro"
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
         Left            =   -66525
         TabIndex        =   175
         Top             =   2640
         Width           =   585
      End
      Begin VB.Label Label60 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Devol"
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
         Left            =   -68325
         TabIndex        =   174
         Top             =   2640
         Width           =   450
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Pagos a Cuenta"
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
         Left            =   -71295
         TabIndex        =   116
         Top             =   5040
         Width           =   1275
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Pagos a Cuenta"
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
         TabIndex        =   114
         Top             =   5040
         Width           =   1275
      End
      Begin VB.Label lblTotalSeguroFarmacia 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Seguros Farmacia"
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
         Left            =   -68820
         TabIndex        =   23
         Top             =   5010
         Width           =   1890
      End
      Begin VB.Label lblTotalSeguroServicio 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Seguros Servicio"
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
         Left            =   6240
         TabIndex        =   21
         Top             =   5040
         Width           =   1815
      End
      Begin VB.Label lblPagoFarmacia 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total a Pagar en Farmacia"
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
         Left            =   -65700
         TabIndex        =   16
         Top             =   5040
         Width           =   2100
      End
      Begin VB.Label lblTotalServicios 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total a Pagar en Servicios"
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
         Left            =   9330
         TabIndex        =   14
         Top             =   5040
         Width           =   2100
      End
   End
   Begin TabDlg.SSTab TabBusqueda 
      Height          =   2415
      Left            =   30
      TabIndex        =   24
      Top             =   510
      Width           =   12600
      _ExtentX        =   22225
      _ExtentY        =   4260
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
      TabCaption(0)   =   "Por Paciente"
      TabPicture(0)   =   "ucEstadoCuenta.ctx":1333
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDatosHistoria"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraDatosAtencion"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmbCtas"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Lista de Pacientes"
      TabPicture(1)   =   "ucEstadoCuenta.ctx":134F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtFiltroApellPat"
      Tab(1).Control(1)=   "btnBuscar"
      Tab(1).Control(2)=   "optHosp"
      Tab(1).Control(3)=   "optEmerg"
      Tab(1).Control(4)=   "optCE"
      Tab(1).Control(5)=   "optPreVentaServ"
      Tab(1).Control(6)=   "optExoneracionesFarmacia"
      Tab(1).Control(7)=   "txtFechaInicio"
      Tab(1).Control(8)=   "txtFechaFin"
      Tab(1).Control(9)=   "optPacientesExternos"
      Tab(1).Control(10)=   "lblFiltroApellPaterno"
      Tab(1).Control(11)=   "Label22"
      Tab(1).Control(12)=   "Label21"
      Tab(1).ControlCount=   13
      Begin VB.ComboBox cmbCtas 
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
         Left            =   4050
         TabIndex        =   6
         Top             =   330
         Width           =   3735
      End
      Begin VB.TextBox txtFiltroApellPat 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -72720
         MaxLength       =   9
         TabIndex        =   126
         Top             =   930
         Width           =   2475
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   -63840
         Picture         =   "ucEstadoCuenta.ctx":136B
         Style           =   1  'Graphical
         TabIndex        =   120
         Top             =   330
         Width           =   1305
      End
      Begin VB.Frame fraDatosAtencion 
         Caption         =   "Datos de paciente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1965
         Left            =   2550
         TabIndex        =   87
         Top             =   360
         Width           =   10005
         Begin VB.TextBox txtDxEgr 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   750
            TabIndex        =   110
            Top             =   1470
            Width           =   4905
         End
         Begin VB.TextBox txtFAltaAdm 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   8685
            TabIndex        =   107
            Top             =   660
            Width           =   1170
         End
         Begin VB.TextBox txtFapertura 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   8685
            TabIndex        =   105
            Top             =   240
            Width           =   1170
         End
         Begin VB.ComboBox cmbFechaIngreso 
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
            Left            =   4590
            TabIndex        =   97
            Top             =   60
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.TextBox txtPaciente 
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
            Left            =   1380
            TabIndex        =   96
            Top             =   270
            Width           =   3255
         End
         Begin VB.ComboBox cmbAgrupar 
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
            ItemData        =   "ucEstadoCuenta.ctx":3FB4
            Left            =   5130
            List            =   "ucEstadoCuenta.ctx":3FB6
            Style           =   2  'Dropdown List
            TabIndex        =   95
            Top             =   60
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox txtFingreso 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6495
            TabIndex        =   94
            Top             =   210
            Width           =   1185
         End
         Begin VB.TextBox txtFegreso 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6495
            TabIndex        =   93
            Top             =   630
            Width           =   1200
         End
         Begin VB.TextBox txtEstadoCuenta 
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
            Left            =   750
            Locked          =   -1  'True
            TabIndex        =   92
            Top             =   660
            Width           =   4920
         End
         Begin VB.TextBox txtCuenta 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   750
            TabIndex        =   91
            Top             =   270
            Width           =   615
         End
         Begin VB.TextBox txtServicio 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   750
            TabIndex        =   90
            Top             =   1080
            Width           =   4905
         End
         Begin VB.TextBox txtDomicilioPacienteEnAtencion 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6510
            TabIndex        =   89
            Top             =   1020
            Width           =   3435
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
            Left            =   4650
            TabIndex        =   88
            Top             =   270
            Width           =   1005
         End
         Begin SISGalenPlus.ucMensajeParpadeando ucMensajeParpadeando1 
            Height          =   405
            Left            =   5700
            TabIndex        =   165
            Top             =   1470
            Visible         =   0   'False
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   714
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            Caption         =   "Dx Egr."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   111
            Top             =   1500
            Width           =   540
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F. Egr.Adm"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7860
            TabIndex        =   108
            Top             =   720
            Width           =   810
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F. Apert.Cta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7740
            TabIndex        =   106
            Top             =   300
            Width           =   915
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nombre"
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
            Left            =   120
            TabIndex        =   104
            Top             =   300
            Width           =   675
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F. Ingreso"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5745
            TabIndex        =   103
            Top             =   270
            Width           =   750
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Agrupar Por"
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
            Left            =   4740
            TabIndex        =   102
            Top             =   90
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F.Alta.Méd"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5700
            TabIndex        =   101
            Top             =   690
            Width           =   795
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta"
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
            Left            =   120
            TabIndex        =   100
            Top             =   720
            Width           =   600
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Serv.Eg"
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
            Left            =   120
            TabIndex        =   99
            Top             =   1110
            Width           =   660
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5880
            TabIndex        =   98
            Top             =   1080
            Width           =   600
         End
      End
      Begin VB.Frame fraDatosHistoria 
         Caption         =   "Búsqueda"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1965
         Left            =   90
         TabIndex        =   82
         Top             =   360
         Width           =   2445
         Begin VB.TextBox txtDctoExoneracionFarm 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1380
            MaxLength       =   20
            TabIndex        =   3
            Top             =   960
            Width           =   1035
         End
         Begin VB.CommandButton btnLeerProductos 
            Height          =   375
            Left            =   1200
            Picture         =   "ucEstadoCuenta.ctx":3FB8
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   1410
            Width           =   1125
         End
         Begin VB.ComboBox cmbIdTipoGenHistoriaClinica 
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
            Left            =   570
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1470
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.TextBox txtNroCuenta 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   960
            MaxLength       =   9
            TabIndex        =   0
            Top             =   180
            Width           =   1005
         End
         Begin VB.CommandButton cmdBuscaCuentaPorApellidos 
            Caption         =   "..."
            Height          =   315
            Left            =   2010
            TabIndex        =   1
            ToolTipText     =   "Busca Cuenta por Apellidos y Nombres"
            Top             =   210
            Width           =   315
         End
         Begin VB.TextBox txtNroOrdenPagoS 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1380
            MaxLength       =   9
            TabIndex        =   2
            Top             =   570
            Width           =   1035
         End
         Begin VB.Label Label48 
            BackStyle       =   0  'Transparent
            Caption         =   "N°Dcto.Exo.Farm"
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
            Left            =   90
            TabIndex        =   112
            Top             =   1020
            Width           =   1395
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
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
            Left            =   30
            TabIndex        =   86
            Top             =   1560
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "N° Historia"
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
            Left            =   330
            TabIndex        =   85
            Top             =   1560
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "N° Cuenta"
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
            Left            =   90
            TabIndex        =   84
            Top             =   270
            Width           =   795
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "N°Ord.Pago Serv"
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
            Left            =   90
            TabIndex        =   83
            Top             =   630
            Width           =   1275
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "TOTAL SERVICIOS"
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
         Height          =   2085
         Left            =   -74790
         TabIndex        =   68
         Top             =   750
         Visible         =   0   'False
         Width           =   3420
         Begin VB.TextBox Text25 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1980
            TabIndex        =   72
            Top             =   270
            Width           =   1215
         End
         Begin VB.TextBox Text24 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1980
            TabIndex        =   71
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox Text23 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1980
            TabIndex        =   70
            Top             =   1170
            Width           =   1215
         End
         Begin VB.TextBox Text22 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1980
            TabIndex        =   69
            Top             =   1590
            Width           =   1215
         End
         Begin VB.Label Label41 
            Caption         =   "INGRESADO (S/.)"
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
            Left            =   60
            TabIndex        =   76
            Top             =   360
            Width           =   1515
         End
         Begin VB.Label Label40 
            Caption         =   "PEND. PAGO (S/.)"
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
            Left            =   60
            TabIndex        =   75
            Top             =   810
            Width           =   1545
         End
         Begin VB.Label Label39 
            Caption         =   "EXONERADO (S/.)"
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
            Left            =   60
            TabIndex        =   74
            Top             =   1230
            Width           =   1605
         End
         Begin VB.Label Label38 
            Caption         =   "DCTOS (S/.)"
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
            Left            =   60
            TabIndex        =   73
            Top             =   1650
            Width           =   1605
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "TOTAL BIENES E INSUMOS"
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
         Height          =   1725
         Left            =   -74760
         TabIndex        =   61
         Top             =   2835
         Visible         =   0   'False
         Width           =   3405
         Begin VB.TextBox Text21 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1980
            TabIndex        =   64
            Top             =   1170
            Width           =   1215
         End
         Begin VB.TextBox Text20 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1980
            TabIndex        =   63
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox Text19 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1980
            TabIndex        =   62
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label37 
            Caption         =   "EXONERADO (S/.)"
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
            Left            =   60
            TabIndex        =   67
            Top             =   1230
            Width           =   1605
         End
         Begin VB.Label Label36 
            Caption         =   "PEND. PAGO (S/.)"
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
            Left            =   60
            TabIndex        =   66
            Top             =   810
            Width           =   1545
         End
         Begin VB.Label Label35 
            Caption         =   "INGRESADO (S/.)"
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
            Left            =   60
            TabIndex        =   65
            Top             =   360
            Width           =   1515
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "TOTALES POR PAGAR"
         Enabled         =   0   'False
         Height          =   2265
         Left            =   -66360
         TabIndex        =   52
         Top             =   810
         Visible         =   0   'False
         Width           =   3405
         Begin VB.TextBox Text18 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1980
            TabIndex        =   56
            Top             =   1230
            Width           =   1215
         End
         Begin VB.TextBox Text17 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1980
            TabIndex        =   55
            Top             =   330
            Width           =   1215
         End
         Begin VB.TextBox Text16 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1980
            TabIndex        =   54
            Top             =   780
            Width           =   1215
         End
         Begin VB.TextBox Text15 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1980
            TabIndex        =   53
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label34 
            Caption         =   "PAGO A CUENTA"
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
            Left            =   210
            TabIndex        =   60
            Top             =   1290
            Width           =   1425
         End
         Begin VB.Label Label33 
            Caption         =   "TOTAL SERVICIOS (S/.)"
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
            Left            =   180
            TabIndex        =   59
            Top             =   390
            Width           =   1785
         End
         Begin VB.Label Label32 
            Caption         =   "TOTAL BIENES (S/.)"
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
            Left            =   180
            TabIndex        =   58
            Top             =   840
            Width           =   1425
         End
         Begin VB.Label Label31 
            Caption         =   "TOTAL"
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
            Left            =   210
            TabIndex        =   57
            Top             =   1710
            Width           =   1425
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "TOTAL PAGADO"
         Enabled         =   0   'False
         Height          =   1845
         Left            =   -66240
         TabIndex        =   45
         Top             =   3180
         Visible         =   0   'False
         Width           =   3405
         Begin VB.TextBox Text14 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1965
            TabIndex        =   48
            Top             =   1260
            Width           =   1215
         End
         Begin VB.TextBox Text13 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1980
            TabIndex        =   47
            Top             =   780
            Width           =   1215
         End
         Begin VB.TextBox Text12 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1980
            TabIndex        =   46
            Top             =   330
            Width           =   1215
         End
         Begin VB.Label Label30 
            Caption         =   "TOTAL"
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
            Left            =   195
            TabIndex        =   51
            Top             =   1290
            Width           =   1425
         End
         Begin VB.Label Label29 
            Caption         =   "TOTAL BIENES (S/.)"
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
            Left            =   180
            TabIndex        =   50
            Top             =   840
            Width           =   1425
         End
         Begin VB.Label Label28 
            Caption         =   "TOTAL SERVICIOS (S/.)"
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
            Left            =   180
            TabIndex        =   49
            Top             =   390
            Width           =   1785
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "ALTA ADMINISTRATIVA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4965
         Left            =   -74880
         TabIndex        =   34
         Top             =   750
         Width           =   8655
         Begin VB.CommandButton Command8 
            Caption         =   "Cerrar Cuenta"
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
            Height          =   525
            Left            =   150
            TabIndex        =   44
            Top             =   3195
            Width           =   1995
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Abrir Cuenta"
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
            Height          =   525
            Left            =   180
            TabIndex        =   43
            Top             =   2205
            Width           =   1995
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Cuenta Pagada"
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
            Height          =   525
            Left            =   120
            TabIndex        =   42
            Top             =   210
            Width           =   1995
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Cuenta Anulada"
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
            Height          =   525
            Left            =   120
            TabIndex        =   41
            Top             =   4200
            Width           =   1995
         End
         Begin VB.TextBox Text11 
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   735
            Left            =   2190
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            TabIndex        =   40
            Text            =   "ucEstadoCuenta.ctx":4418
            Top             =   210
            Width           =   6375
         End
         Begin VB.TextBox Text10 
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   735
            Left            =   2250
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            TabIndex        =   39
            Text            =   "ucEstadoCuenta.ctx":441E
            Top             =   2220
            Width           =   6375
         End
         Begin VB.TextBox Text9 
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   735
            Left            =   2220
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            TabIndex        =   38
            Text            =   "ucEstadoCuenta.ctx":4424
            Top             =   3225
            Width           =   6375
         End
         Begin VB.TextBox Text8 
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   705
            Left            =   2190
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            TabIndex        =   37
            Text            =   "ucEstadoCuenta.ctx":442A
            Top             =   4230
            Width           =   6375
         End
         Begin VB.TextBox Text7 
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   735
            Left            =   2220
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            TabIndex        =   36
            Text            =   "ucEstadoCuenta.ctx":4430
            Top             =   1215
            Width           =   6375
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Pendiente Pago Seguros"
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
            Height          =   525
            Left            =   150
            TabIndex        =   35
            Top             =   1200
            Width           =   1995
         End
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   -63510
         TabIndex        =   33
         Top             =   5580
         Width           =   945
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Todos/Ninguno"
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
         Left            =   -74850
         TabIndex        =   32
         Top             =   5580
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -73410
         Picture         =   "ucEstadoCuenta.ctx":4436
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   5550
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.Frame Frame7 
         Caption         =   "RECALCULO de Cuenta de Atención"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4965
         Left            =   -66120
         TabIndex        =   27
         Top             =   750
         Width           =   3495
         Begin VB.CommandButton Command1 
            Caption         =   "Cambia a otro PLAN DE ATENCION"
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
            Height          =   525
            Left            =   120
            TabIndex        =   29
            Top             =   2280
            Width           =   3225
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   1395
            Left            =   120
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            TabIndex        =   28
            Text            =   "ucEstadoCuenta.ctx":4C67
            Top             =   300
            Width           =   3195
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   315
            Left            =   120
            TabIndex        =   30
            Top             =   1800
            Width           =   3225
            _ExtentX        =   5689
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   "DataCombo1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   -68730
         TabIndex        =   26
         Text            =   "0"
         Top             =   5550
         Width           =   1005
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   -63900
         TabIndex        =   25
         Top             =   2580
         Width           =   945
      End
      Begin UltraGrid.SSUltraGrid SSUltraGrid1 
         Height          =   2085
         Left            =   -74910
         TabIndex        =   77
         Top             =   420
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   3678
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
         Caption         =   "Resumen por Punto de Carga"
      End
      Begin UltraGrid.SSUltraGrid SSUltraGrid2 
         Height          =   2805
         Left            =   -74910
         TabIndex        =   78
         Top             =   3000
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   4948
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
         Caption         =   "...."
      End
      Begin Threed.SSOption optHosp 
         Height          =   255
         Left            =   -74835
         TabIndex        =   121
         Top             =   630
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   450
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
         Caption         =   "Hospitalizados"
         Value           =   -1
      End
      Begin Threed.SSOption optEmerg 
         Height          =   255
         Left            =   -73080
         TabIndex        =   122
         Top             =   630
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   450
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
         Caption         =   "Emergencia"
      End
      Begin Threed.SSOption optCE 
         Height          =   255
         Left            =   -71565
         TabIndex        =   123
         Top             =   630
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
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
         Caption         =   "Consulta Externa"
      End
      Begin Threed.SSOption optPreVentaServ 
         Height          =   255
         Left            =   -69495
         TabIndex        =   124
         Top             =   630
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   450
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
         Caption         =   "PreVentas Servicios"
      End
      Begin Threed.SSOption optExoneracionesFarmacia 
         Height          =   255
         Left            =   -67260
         TabIndex        =   125
         Top             =   630
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   450
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
         Caption         =   "Exoneraciones en Farmacia"
      End
      Begin MSMask.MaskEdBox txtFechaInicio 
         Height          =   315
         Left            =   -67770
         TabIndex        =   117
         Top             =   330
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaFin 
         Height          =   315
         Left            =   -65670
         TabIndex        =   118
         Top             =   330
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin Threed.SSOption optPacientesExternos 
         Height          =   255
         Left            =   -64395
         TabIndex        =   131
         Top             =   630
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
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
         Caption         =   "Pacientes Externos"
      End
      Begin VB.Label lblFiltroApellPaterno 
         Caption         =   "Filtra por Apellido Paterno"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -74850
         TabIndex        =   128
         Top             =   930
         Width           =   2205
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "hasta"
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
         Left            =   -66225
         TabIndex        =   127
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label21 
         Caption         =   "F.Ingreso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -68580
         TabIndex        =   119
         Top             =   360
         Width           =   765
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total a Pagar en Farmacia"
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
         Left            =   -65850
         TabIndex        =   81
         Top             =   5640
         Width           =   2280
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Seguros Farmacia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -70440
         TabIndex        =   80
         Top             =   5610
         Width           =   1680
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -64320
         TabIndex        =   79
         Top             =   2640
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdExoneracion 
      Caption         =   "Imprimir EXONERACION"
      Enabled         =   0   'False
      Height          =   1005
      Left            =   5565
      Picture         =   "ucEstadoCuenta.ctx":4C6D
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8460
      Width           =   1695
   End
   Begin VB.CommandButton bntLiquidacion 
      Caption         =   "Imprimir LIQUIDACION"
      Enabled         =   0   'False
      Height          =   1005
      Left            =   3755
      Picture         =   "ucEstadoCuenta.ctx":5146
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8460
      Width           =   1695
   End
   Begin VB.CommandButton cmdImprimeCtaPorServicioHosp 
      Caption         =   "Imprimir ESTADO DE CUENTA por Servicio Hosp"
      Enabled         =   0   'False
      Height          =   1005
      Left            =   1915
      Picture         =   "ucEstadoCuenta.ctx":561F
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8460
      Width           =   1725
   End
   Begin VB.CommandButton btnImprimir 
      Caption         =   "Imprimir ESTADO DE CUENTA por Pto.Carga"
      Enabled         =   0   'False
      Height          =   1005
      Left            =   75
      Picture         =   "ucEstadoCuenta.ctx":5AF8
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8460
      Width           =   1725
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Estado de cuenta del paciente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   12660
   End
   Begin VB.Menu mnuExoneracion 
      Caption         =   "Exoneracion"
      Begin VB.Menu mnuAgregarExoneracion 
         Caption         =   "Agregar Exoneracion"
      End
   End
   Begin VB.Menu mnuACuenta 
      Caption         =   "Pagos A Cuenta"
      Begin VB.Menu mnuAgregaACuenta 
         Caption         =   "Agregar Pagos A Cuenta"
      End
   End
   Begin VB.Menu mnuPolizas 
      Caption         =   "Polizas de Seguro"
      Begin VB.Menu mnuAgregaPoliza 
         Caption         =   "Agrega Poliza"
      End
   End
   Begin VB.Menu mnuBienes 
      Caption         =   "Bienes"
      Begin VB.Menu mnuAgregaBienes 
         Caption         =   "Agregar Bienes"
      End
   End
   Begin VB.Menu mnuServicios 
      Caption         =   "Servicios"
      Begin VB.Menu mnuAgregaServicios 
         Caption         =   "Agregar Servicios"
      End
   End
End
Attribute VB_Name = "ucEstadoCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para mostrar ESTADO DE CUENTA de un Paciente
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_ReglasSISgalenhos As New SIGHSis.ReglasSISgalenhos
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasReportes As New ReglasReportes
Dim mo_PermisosFacturacion As New PermisosFacturacion
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_cmbIdTipoGenHistoriaClinica As New sighentidades.ListaDespleglable
Dim mo_cmbFechaIngreso As New sighentidades.ListaDespleglable
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes

Dim ml_idPaciente  As Long
Dim mo_DoAtencion As New DOAtencion
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_DOCuentaAtencion As New DOCuentaAtencion
Dim md_Total As Double
Dim md_TotalBien As Double
Dim md_TotalServ As Double
Dim md_TotalPagado As Double
Dim md_TotalBienPagado As Double
Dim ml_idCuentaAtencion As Long
Dim ml_idAtencion As Long
Dim ml_idUsuarioConPermisoEnSISoEXOoSOAT As Long
Dim ml_idUsuario As Long
Dim mo_ReporteUtil As New sighentidades.ReporteUtil
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_IdTipoServicio As sghTipoServicio
Dim oRsFuentesFinanciamiento As New ADODB.Recordset
Dim oRsCuentaCabecera As New Recordset
Dim oRsCuentaDetalle As New Recordset
Dim oRsCuentasPorTipoServicio As New Recordset
Dim oRsReembolsos As New Recordset
Dim oRsFormaPago As New Recordset
Dim rsReporte As New Recordset
Dim rsItemsDonaciones As New Recordset
Dim lnIdPlanActual As Long
Dim lnIdTipoFinanciamientoActual As Long
Dim ml_idEstadoCuentaAtencion As sghEstadoCuenta
Dim lc_TipoFinanciamientoPermitidos As String
Dim gridInfra As New GridInfragistic
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim lcdTipoFinanciamiento As String
Dim lbTieneAccesoActualizar As Boolean
Dim lnIdPagosACuenta As Long
Dim lnTotalPagosAdelantados As Double
Dim lnEstadoFacturacionAtendidoOpreventa As sghEstadoFacturacion
Dim lcListaDeOrdenesDePago As String
Dim ml_lbEsPacienteExterno As Boolean
Dim lnPagosXdevoluciones As Double
Dim lnIdPagosXdevoluciones As Long
Dim lcSql As String
Dim oConexionConsulta As New Connection
Dim ml_dCondicionAlta As String
Dim ml_idTipoSexo As Long
Dim ml_lbPuedeVerResultados As Boolean
Dim lcHoraInicioProceso As String, lcHoraFinalProceso As String, mb_ProcesoEnElServidor As Boolean, lbGeneraReciboPago As Boolean
Dim ml_lnHWnd As Long
Property Let lnHWnd(lValue As Long)
    ml_lnHWnd = lValue
End Property
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property


'***************daniel barrantes**************
'***************Impresion de Liquidacion SOAT
'***************
Private Sub bntLiquidacion_Click()
    Dim iFila As Long: Dim iCol As Integer
    Dim rsReporte As New Recordset
    Dim ms_EstadosFacturacion As String
    Dim ms_TiposFinanciamiento As String
    Dim ml_AgruparPor As Long
    Dim idPuntoCarga As Long: Dim lnIdTipoServicio As Long
    Dim lcEstancia As String
    
    Dim lnTSubTotal As Double: Dim lnTSubTotalSIS As Double: Dim lnTSubTotalEXO As Double
    Dim lnTSubTotalPAGO As Double: Dim lnTSubTotalDEBE As Double: Dim lnTSubTotalSALDO As Double
    
    Dim lnTotal As Double: Dim lnTotalSIS As Double: Dim lnTotalEXO As Double
    Dim lnTotalPAGO As Double: Dim lnTotalDEBE As Double: Dim lnTotalSALDO As Double
    
    Dim lnDebe As Double: Dim lnPago As Double: Dim lnSaldo As Double
    Dim lnTDebe As Double: Dim lnTPago As Double: Dim lnTSaldo As Double
    Dim lnSIS As Double: Dim lnEXO As Double: Dim lnTotalCredito  As Double: Dim lnSOAT As Double
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim CantidadSOAT As Long: Dim PrecioSOAT As Double
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
        Dim lnHWnd As Long
    Else
        Dim oExcel As Excel.Application
        Dim oWorkBookPlantilla As Workbook
        Dim oWorkBook As Workbook
        Dim oWorkSheet As Worksheet
        Dim oRange As range
        Dim range As Excel.range
        Dim borders As Excel.borders
    End If
    
    
    If txtPaciente.Text = "" Then
        MsgBox "Tiene que LEER el Estado de Cuenta del Paciente", vbInformation, "Resultado"
    Else
        
        
        MousePointer = 11
        If lbEsOpenOffice = True Then
            'Abre el archivo ExcelOpenOffice
            lcArchivoExcel = App.Path + "\Plantillas\ELiquidacion.ods"
    '        FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
    '        Chemin = "file:///" & App.Path & "\Plantillas\"
    '        Chemin = Replace(Chemin, "\", "/")
    '        Fichier = Chemin & "/OpenOffice.ods"
            '
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
            'Encabezado de Pagina
            mo_CabeceraReportes.CabeceraReportes Document, True
            ' Pone la ventana en primer plano, pasándole el Hwnd
            ret = SetForegroundWindow(lnHWnd)
        Else
            'Crea nueva hoja
            Set oExcel = GalenhosExcelApplication()  'New Excel.Application
            Set oWorkBook = oExcel.Workbooks.Add
            'Abre, copia y cierra la plantilla
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\ELiquidacion.xls")
            oWorkBookPlantilla.Worksheets("Liquidacion").Copy Before:=oWorkBook.Sheets(1)
            oWorkBookPlantilla.Close
            'Activa la primera hoja
            Set oWorkSheet = oWorkBook.Sheets(1)
            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
        End If
        '*******************************************Inicio de Impresion
        'Atencion
        Set rsReporte = mo_ReglasFacturacion.AtencionesFiltraEgresosPorIdAtencion(ml_idAtencion)
        If rsReporte.RecordCount > 0 Then
            lcEstancia = "Del " & mo_ReporteUtil.NullToVacio(rsReporte.Fields!FechaIngreso) & " " & _
                         " - " & mo_ReporteUtil.NullToVacio(rsReporte.Fields!HoraIngreso) & " al " & _
                         mo_ReporteUtil.NullToVacio(rsReporte.Fields!FechaEgreso) & " " & _
                         mo_ReporteUtil.NullToVacio(rsReporte.Fields!HoraEgreso)
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(2, 2).setFormula(txtPaciente.Text)
                Call Feuille.getcellbyposition(5, 2).setFormula(Trim(txtNroHistoria.Text) & "        Cuenta: " & txtCuenta.Text)
                If rsReporte.Fields!idTipoServicio = sighentidades.sghHospitalizacion Then
                   Call Feuille.getcellbyposition(5, 10).setFormula(lcEstancia)
                ElseIf rsReporte.Fields!idTipoServicio = sighentidades.sghConsultaExterna Then
                   Call Feuille.getcellbyposition(2, 11).setFormula(lcEstancia)
                Else
                   Call Feuille.getcellbyposition(2, 10).setFormula(lcEstancia)
                End If
            Else
                oWorkSheet.Cells(3, 3).Value = txtPaciente.Text
                oWorkSheet.Cells(3, 6).Value = Trim(txtNroHistoria.Text) & "        Cuenta: " & txtCuenta.Text
                If rsReporte.Fields!idTipoServicio = sighentidades.sghHospitalizacion Then
                   oWorkSheet.Cells(11, 6).Value = lcEstancia
                ElseIf rsReporte.Fields!idTipoServicio = sighentidades.sghConsultaExterna Then
                   oWorkSheet.Cells(12, 3).Value = lcEstancia
                Else
                   oWorkSheet.Cells(11, 3).Value = lcEstancia
                End If
            End If
        End If
        rsReporte.Close
        'Diagnosticos
        Set rsReporte = mo_ReglasFacturacion.AtencionesFiltraDiagnosticosPorIdAtencion(ml_idAtencion, lnIdTipoServicio)
        If rsReporte.RecordCount > 0 Then
           iFila = 6
           rsReporte.MoveFirst
           Do While Not rsReporte.EOF
            If lbEsOpenOffice = True Then
               Call Feuille.getcellbyposition(2, iFila - 1).setFormula(Left(rsReporte.Fields!dTipo, 1) & " " & rsReporte.Fields!Cie10 & " " & rsReporte.Fields!dDiagnostico)
            Else
               oWorkSheet.Cells(iFila, 3).Value = Left(rsReporte.Fields!dTipo, 1) & " " & rsReporte.Fields!Cie10 & " " & rsReporte.Fields!dDiagnostico
            End If
            iFila = iFila + 1
            If iFila > 8 Then
               Exit Do
            Else
               rsReporte.MoveNext
            End If
           Loop
        End If
        rsReporte.Close
        
        iFila = 15
        iCol = 2
        If lbEsOpenOffice = True Then
            Select Case ml_idUsuarioConPermisoEnSISoEXOoSOAT
            Case 2   'SIS
                Call Feuille.getcellbyposition(4, iFila - 2).setFormula("Precio SIS")
            Case 3   'SOAT
                Call Feuille.getcellbyposition(4, iFila - 2).setFormula("Precio SOAT")
            Case 4   'CONVENIO
                Call Feuille.getcellbyposition(4, iFila - 2).setFormula("Precio CONVENIO")
            Case 9   'EXONERACIONES
                Call Feuille.getcellbyposition(5, iFila - 2).setFormula("Imp.Exonerado")
            End Select
        Else
            Select Case ml_idUsuarioConPermisoEnSISoEXOoSOAT
            Case 2   'SIS
                oWorkSheet.Cells(iFila - 1, 5).Value = "Precio SIS"
            Case 3   'SOAT
                oWorkSheet.Cells(iFila - 1, 5).Value = "Precio SOAT"
            Case 4   'CONVENIO
                oWorkSheet.Cells(iFila - 1, 5).Value = "Precio CONVENIO"
            Case 9   'EXONERACIONES
                oWorkSheet.Cells(iFila - 1, 6).Value = "Imp.Exonerado"
            End Select
        End If
        lnTotal = 0: lnTotalSIS = 0: lnTotalEXO = 0: lnTotalPAGO = 0: lnTotalDEBE = 0: lnTotalSALDO = 0: lnTotalCredito = 0
        lnTDebe = 0: lnTPago = 0: lnTSaldo = 0
        'Farmacia
        Set rsReporte = ucFacturacionBienes.FacturacionProductos
        If rsReporte.RecordCount > 0 Then
            rsReporte.MoveFirst
            If lbEsOpenOffice = True Then
               Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Farmacia")
            Else
               oWorkSheet.Cells(iFila, 2).Value = "Farmacia"
            End If
            iFila = iFila + 1
            lnTSubTotal = 0: lnTSubTotalSIS = 0: lnTSubTotalEXO = 0
            lnTSubTotalPAGO = 0: lnTSubTotalDEBE = 0: lnTSubTotalSALDO = 0
            Do While Not rsReporte.EOF
                    Select Case ml_idUsuarioConPermisoEnSISoEXOoSOAT
                    Case 2   'SIS
                        lnSOAT = IIf(IsNull(rsReporte.Fields!ImporteSIS), 0, rsReporte.Fields!ImporteSIS)
                        CantidadSOAT = IIf(IsNull(rsReporte.Fields("CantidadSIS").Value), 0, rsReporte.Fields("CantidadSIS").Value)
                        PrecioSOAT = IIf(IsNull(rsReporte.Fields("PrecioSIS").Value), 0, rsReporte.Fields("PrecioSIS").Value)
                    Case 3   'SOAT
                        lnSOAT = IIf(IsNull(rsReporte.Fields!ImporteSOAT), 0, rsReporte.Fields!ImporteSOAT)
                        CantidadSOAT = IIf(IsNull(rsReporte.Fields("CantidadSOAT").Value), 0, rsReporte.Fields("CantidadSOAT").Value)
                        PrecioSOAT = IIf(IsNull(rsReporte.Fields("PrecioSOAT").Value), 0, rsReporte.Fields("PrecioSOAT").Value)
                    Case 4   'CONVENIO
                            lnSOAT = IIf(IsNull(rsReporte.Fields!ImporteConv), 0, rsReporte.Fields!ImporteConv)
                            CantidadSOAT = IIf(IsNull(rsReporte.Fields("CantidadConv").Value), 0, rsReporte.Fields("CantidadConv").Value)
                            PrecioSOAT = IIf(IsNull(rsReporte.Fields("precioConv").Value), 0, rsReporte.Fields("precioConv").Value)
                    Case 9   'EXONERACIONES
                        
                        lnSOAT = IIf(IsNull(rsReporte.Fields!importeEXO), 0, rsReporte.Fields!importeEXO)
                        CantidadSOAT = 0
                        PrecioSOAT = 0
                        
                    End Select
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(1, iFila - 1).setFormula(rsReporte.Fields("Codigo").Value)
                        Call Feuille.getcellbyposition(2, iFila - 1).setFormula(rsReporte.Fields("NombreProducto").Value)
                        If lnSOAT > 0 Then
                            Call Feuille.getcellbyposition(3, iFila - 1).setFormula(CantidadSOAT)
                            Call Feuille.getcellbyposition(4, iFila - 1).setFormula(Format(PrecioSOAT, "####,###.#0"))
                            Call Feuille.getcellbyposition(5, iFila - 1).setFormula(Format(lnSOAT, "####,###.#0"))
                        Else
                            Call Feuille.getcellbyposition(4, iFila - 1).setFormula("0")
                            Call Feuille.getcellbyposition(4, iFila - 1).setFormula("0")
                        End If
                        Call Feuille.getcellbyposition(2, iFila - 1).setFormula(rsReporte.Fields("NombreProducto").Value)
                    Else
                        oWorkSheet.Cells(iFila, 2).Value = rsReporte.Fields("Codigo").Value
                        oWorkSheet.Cells(iFila, 3).Value = rsReporte.Fields("NombreProducto").Value
                        If lnSOAT > 0 Then
                            oWorkSheet.Cells(iFila, 4).Value = CantidadSOAT
                            oWorkSheet.Cells(iFila, 5).Value = Format(PrecioSOAT, "####,###.#0")
                            oWorkSheet.Cells(iFila, 6).Value = Format(lnSOAT, "####,###.#0")
                        Else
                            oWorkSheet.Cells(iFila, 5).Value = 0
                            oWorkSheet.Cells(iFila, 6).Value = 0
                        End If
                        oWorkSheet.Cells(iFila, 8).Value = rsReporte.Fields!IdOrden
                    End If
                    lnTSubTotal = lnTSubTotal + lnSOAT
                    lnTotal = lnTotal + lnSOAT
                    
                    iFila = iFila + 1
                rsReporte.MoveNext
             Loop
             If lbEsOpenOffice = True Then
                Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(3) & CStr(iFila) & ":" & mo_ReglasReportes.BuscaNombreColumna(8) & CStr(iFila))
                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                Call Feuille.getcellbyposition(6, iFila - 1).setFormula("0")
             Else
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 3, iFila, 8
                oWorkSheet.Cells(iFila, 7).Value = Format(lnTSubTotal, "####,###.#0")
             End If
             iFila = iFila + 1
        End If
        rsReporte.Close
        'Servicios
        Set rsReporte = ucFacturacionServicios.FacturacionProductos
        If rsReporte.RecordCount > 0 Then
            rsReporte.Sort = "IdPuntoCarga"
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF
                idPuntoCarga = rsReporte.Fields("IdPuntoCarga").Value
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(1, iFila - 1).setFormula(mo_ReglasComunes.FactPuntosCargaSeleccionarPorIdDevDescripcion(rsReporte.Fields("IdPuntoCarga").Value))
                Else
                    oWorkSheet.Cells(iFila, 2).Value = mo_ReglasComunes.FactPuntosCargaSeleccionarPorIdDevDescripcion(rsReporte.Fields("IdPuntoCarga").Value)
                    'oWorkSheet.Cells(iFila, 2).Value = FactPuntosCargaSeleccionarPorId(rsReporte.Fields("IdPuntoCarga").Value)
                End If
                iFila = iFila + 1
                lnTSubTotal = 0: lnTSubTotalSIS = 0: lnTSubTotalEXO = 0
                lnTSubTotalPAGO = 0: lnTSubTotalDEBE = 0: lnTSubTotalSALDO = 0
                Do While Not rsReporte.EOF And idPuntoCarga = rsReporte.Fields("IdPuntoCarga").Value
                        Select Case ml_idUsuarioConPermisoEnSISoEXOoSOAT
                        Case 2   'SIS
                            lnSOAT = IIf(IsNull(rsReporte.Fields!ImporteSIS), 0, rsReporte.Fields!ImporteSIS)
                            CantidadSOAT = IIf(IsNull(rsReporte.Fields("CantidadSIS").Value), 0, rsReporte.Fields("CantidadSIS").Value)
                            PrecioSOAT = IIf(IsNull(rsReporte.Fields("PrecioSIS").Value), 0, rsReporte.Fields("PrecioSIS").Value)
                        Case 3   'SOAT
                            lnSOAT = IIf(IsNull(rsReporte.Fields!ImporteSOAT), 0, rsReporte.Fields!ImporteSOAT)
                            CantidadSOAT = IIf(IsNull(rsReporte.Fields("CantidadSOAT").Value), 0, rsReporte.Fields("CantidadSOAT").Value)
                            PrecioSOAT = IIf(IsNull(rsReporte.Fields("PrecioSOAT").Value), 0, rsReporte.Fields("PrecioSOAT").Value)
                        Case 4   'CONVENIO
                                lnSOAT = IIf(IsNull(rsReporte.Fields!ImporteConv), 0, rsReporte.Fields!ImporteConv)
                                CantidadSOAT = IIf(IsNull(rsReporte.Fields("CantidadConv").Value), 0, rsReporte.Fields("CantidadConv").Value)
                                PrecioSOAT = IIf(IsNull(rsReporte.Fields("precioConv").Value), 0, rsReporte.Fields("precioConv").Value)
                        Case 9   'EXONERACIONES
                            lnSOAT = IIf(IsNull(rsReporte.Fields!importeEXO), 0, rsReporte.Fields!importeEXO)
                            CantidadSOAT = 0
                            PrecioSOAT = 0
                        End Select
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(1, iFila - 1).setFormula(rsReporte.Fields("Codigo").Value)
                            Call Feuille.getcellbyposition(2, iFila - 1).setFormula(rsReporte.Fields("NombreProducto").Value)
                            Call Feuille.getcellbyposition(3, iFila - 1).setFormula(CantidadSOAT)
                            If lnSOAT > 0 Then
                                Call Feuille.getcellbyposition(4, iFila - 1).setFormula(Format(PrecioSOAT, "####,###.#0"))
                                Call Feuille.getcellbyposition(5, iFila - 1).setFormula(Format(lnSOAT, "####,###.#0"))
                            Else
                                Call Feuille.getcellbyposition(4, iFila - 1).setFormula("0")
                                Call Feuille.getcellbyposition(5, iFila - 1).setFormula("0")
                            End If
                            Call Feuille.getcellbyposition(7, iFila - 1).setFormula(rsReporte.Fields!IdOrden)
                           
                        Else
                            oWorkSheet.Cells(iFila, 2).Value = rsReporte.Fields("Codigo").Value
                            oWorkSheet.Cells(iFila, 3).Value = rsReporte.Fields("NombreProducto").Value
                            oWorkSheet.Cells(iFila, 4).Value = CantidadSOAT
                            If lnSOAT > 0 Then
                                oWorkSheet.Cells(iFila, 5).Value = Format(PrecioSOAT, "####,###.#0")
                                oWorkSheet.Cells(iFila, 6).Value = Format(lnSOAT, "####,###.#0")
                            Else
                                oWorkSheet.Cells(iFila, 5).Value = 0
                                oWorkSheet.Cells(iFila, 6).Value = 0
                            End If
                            oWorkSheet.Cells(iFila, 8).Value = rsReporte.Fields!IdOrden
                        End If
                        lnTSubTotal = lnTSubTotal + lnSOAT
                        lnTotal = lnTotal + lnSOAT
                        
                        iFila = iFila + 1
                    rsReporte.MoveNext
                    If rsReporte.EOF Then
                       Exit Do
                    End If
                Loop
                If lbEsOpenOffice = True Then
                    Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(3) & CStr(iFila) & ":" & mo_ReglasReportes.BuscaNombreColumna(8) & CStr(iFila))
                    mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                    Call Feuille.getcellbyposition(6, iFila - 1).setFormula(Format(lnTSubTotal, "####,###.#0"))
                Else
                    mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 3, iFila, 8
                    oWorkSheet.Cells(iFila, 7).Value = Format(lnTSubTotal, "####,###.#0")
                End If
                iFila = iFila + 1
             Loop
        End If
        iFila = iFila + 1
        If lbEsOpenOffice = True Then
            Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(2) & CStr(iFila) & ":" & mo_ReglasReportes.BuscaNombreColumna(8) & CStr(iFila))
            mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
            Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Total: ")
            Call Feuille.getcellbyposition(6, iFila - 1).setFormula(Format(lnTotal, "####,###.#0"))
        Else
            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 8
            oWorkSheet.Cells(iFila, 2).Value = "Total: "
            oWorkSheet.Cells(iFila, 7).Value = Format(lnTotal, "####,###.#0")
        End If
        If lbEsOpenOffice = True Then
            Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
            PrintArea(0).Sheet = 0
            PrintArea(0).startcolumn = 1
            PrintArea(0).StartRow = 0
            PrintArea(0).EndColumn = 14
            PrintArea(0).EndRow = iFila
            Call Feuille.SetPrintAreas(PrintArea())
            Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
            MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
        Else
            oWorkSheet.PageSetup.PrintTitleRows = "$1:$14"
            If oWorkSheet.PageSetup.PrintArea <> "" Then oWorkSheet.PageSetup.PrintArea = "$A$1:$I$" & (iFila + 2) 'sighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
            oExcel.Visible = True
            oWorkSheet.PrintPreview
            'oWorkSheet.PrintOut
        End If
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
        'encabezado de pagina
        Set PageStyles = Nothing
        Set Sheet = Nothing
        Set StyleFamilies = Nothing
        Set DefPage = Nothing
        Set Htext = Nothing
        Set Hcontent = Nothing
    Else
        'liberar memoria
        Set oExcel = Nothing
        Set oWorkBookPlantilla = Nothing
        Set oWorkBook = Nothing
        Set oWorkSheet = Nothing
    End If
    MousePointer = 1
End Sub


Private Sub btnAbrirCuenta_Click()
    If mo_ReglasSISgalenhos.SisFUAyaFueEnviadoAlSisLIMA(ml_idCuentaAtencion, lnIdTipoFinanciamientoActual, wxParametro302) = True Then
       Exit Sub
    End If
    If MsgBox("Esta seguro que desea ABRIR la Cuenta", vbQuestion + vbYesNo, "Facturación") = vbYes Then
        Dim Login As New Login
        Login.UsuarioDeEstadoDeCuenta = ml_idUsuario
        Login.CargaDesdeOtraOpcion = True
        Login.Show vbModal
        If Not Login.Autenticado Or Login.IdUsuarioAutenticado <> ml_idUsuario Then
            Exit Sub
        End If
        mo_DOCuentaAtencion.IdUsuarioAuditoria = ml_idUsuario
           

        If mo_ReglasFacturacion.CuentasAtencionAbrir(mo_DOCuentaAtencion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtEstadoCuenta.Text) Then
            mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar ml_idCuentaAtencion, False, 0
            MsgBox "La Cuenta se ha Abierto correctamente", vbInformation, "Facturación"
            LimparDatos
        Else
            MsgBox "No se pudo Abrir la Cuenta", vbInformation, "Facturación"
        End If
    End If
End Sub

'***************daniel barrantes**************
'***************Grabar Cantidades/precios/Importes registrados por SIS/SOAT/
'***************/ASISTENTA SOCIAL/CONVENIO FOSPOLIS
Private Sub btnAceptar_Click()
    If txtPaciente.Text = "" Then
        MsgBox "Tiene que LEER el Estado de Cuenta del Paciente", vbInformation, "Resultado"
    Else
        Dim oGrabaDatos As New SighFacturacion.dllFactUcEstadoCuenta
        oGrabaDatos.GrabaCantidadesPreciosRegistradosEnSisSoatExoConvenio ml_idUsuarioConPermisoEnSISoEXOoSOAT, ucFacturacionServicios.FacturacionProductos, ucFacturacionBienes.FacturacionProductos, ml_idUsuario, Val(txtCuenta.Text), ml_idPaciente, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lnIdPlanActual, txtEstadoCuenta.Text, lnEstadoFacturacionAtendidoOpreventa
        If lbGeneraReciboPago = True Then
            If MsgBox("   Los datos se grabaron correctamente  " & Chr(13) & _
                      "                                        " & Chr(13) & _
                      "   ¿ Imprime la HOJA DE EXONERACION ?   ", vbQuestion + vbYesNo, "Estado de Cuenta") = vbYes Then
               txtNroCuenta.Text = txtCuenta.Text
               txtNroCuenta_KeyPress 13
               cmdExoneracion_Click
            End If
        End If
        Set oGrabaDatos = Nothing
        mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar ml_idCuentaAtencion, False, 0
        LimparDatos
    End If
End Sub

Private Sub btnBuscar_Click()
    MousePointer = 11
    txtFiltroApellPat.Visible = False
    lblFiltroApellPaterno.Visible = False
    If optPreVentaServ.Value = True Then
       Set oRsCuentasPorTipoServicio = mo_ReglasFacturacion.FactOrdenServicioPreventasServicio(CDate(txtFechaInicio.Text), CDate(txtFechaFin.Text))
    ElseIf optExoneracionesFarmacia.Value = True Then
       Set oRsCuentasPorTipoServicio = mo_ReglasFarmacia.farmMovimientoVentasExoneracionesEnFarmacia(CDate(txtFechaInicio.Text), CDate(txtFechaFin.Text))
    ElseIf optPacientesExternos.Value = True Then
       Set oRsCuentasPorTipoServicio = mo_AdminAdmision.AtencionesSeleccionarPacExtPorFechas1(CDate(txtFechaInicio.Text), CDate(txtFechaFin.Text))
    Else
       Set oRsCuentasPorTipoServicio = mo_ReglasFacturacion.AtencionesSeleccionarPorTipoServicio(IIf(optCE.Value = True, 1, IIf(optEmerg.Value = True, 2, 3)), CDate(txtFechaInicio.Text), CDate(txtFechaFin.Text))
       txtFiltroApellPat.Visible = True
       lblFiltroApellPaterno.Visible = True
    End If
    Set grdCuentasPorTipoServicio.DataSource = oRsCuentasPorTipoServicio
    MousePointer = 1
End Sub

Private Sub btnCerrarCuenta_Click()
    If mo_ReglasSISgalenhos.SisFUAyaFueEnviadoAlSisLIMA(ml_idCuentaAtencion, lnIdTipoFinanciamientoActual, wxParametro302) = True Then
       Exit Sub
    End If
    If MsgBox("Esta seguro que desea CERRAR la Cuenta", vbQuestion + vbYesNo, "Facturación") = vbYes Then
        Dim Login As New Login
        Login.UsuarioDeEstadoDeCuenta = ml_idUsuario
        Login.CargaDesdeOtraOpcion = True
        Login.Show vbModal
        If Not Login.Autenticado Or Login.IdUsuarioAutenticado <> ml_idUsuario Then
            Exit Sub
        End If
        mo_DOCuentaAtencion.IdUsuarioAuditoria = ml_idUsuario
        If txtTotalServicios.Text = "" Then txtTotalServicios.Text = "0"
        If txtTotalFarmacia.Text = "" Then txtTotalFarmacia.Text = "0"
        mo_DOCuentaAtencion.TotalPorPagar = CCur(txtTotalServicios.Text) + CCur(txtTotalFarmacia.Text)
        If mo_ReglasFacturacion.CuentasAtencionCerrar(mo_DOCuentaAtencion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtEstadoCuenta.Text) Then
            mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar ml_idCuentaAtencion, False, 0
            MsgBox "La Cuenta se ha cerrado correctamente", vbInformation, "Facturación"
            LimparDatos
        Else
            MsgBox "No se pudo cerrar la cuenta", vbInformation, "Facturación"
        End If
    End If
End Sub

Private Sub btnCtaAnulada_Click()
    If mo_ReglasSISgalenhos.SisFUAyaFueEnviadoAlSisLIMA(ml_idCuentaAtencion, lnIdTipoFinanciamientoActual, wxParametro302) = True Then
       Exit Sub
    End If
    If MsgBox("Esta seguro que la Cuenta pase a estado=ANULADA", vbQuestion + vbYesNo, "Facturación") = vbYes Then
        Dim Login As New Login
        Login.UsuarioDeEstadoDeCuenta = ml_idUsuario
        Login.CargaDesdeOtraOpcion = True
        Login.Show vbModal
        If Not Login.Autenticado Or Login.IdUsuarioAutenticado <> ml_idUsuario Then
            Exit Sub
        End If
        If txtTotalServicios.Text = "" Then txtTotalServicios.Text = "0"
        If txtTotalFarmacia.Text = "" Then txtTotalFarmacia.Text = "0"
        mo_DOCuentaAtencion.IdUsuarioAuditoria = ml_idUsuario
        mo_DOCuentaAtencion.TotalPorPagar = Val(txtTotalServicios.Text) + Val(txtTotalFarmacia.Text)
        If mo_ReglasFacturacion.CuentasAtencionAnulada(mo_DOCuentaAtencion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtEstadoCuenta.Text) = True Then
            mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar ml_idCuentaAtencion, False, 0
            MsgBox "La Cuenta pasó a estado=ANULADA", vbInformation, "Facturación"
            LimparDatos
        Else
            MsgBox "No se pudo", vbInformation, "Facturación"
        End If
    End If

End Sub

Private Sub btnCtaGarante_Click()
  If mo_ReglasSISgalenhos.SisFUAyaFueEnviadoAlSisLIMA(ml_idCuentaAtencion, lnIdTipoFinanciamientoActual, wxParametro302) = True Then
       Exit Sub
  End If
  If MsgBox("Esta seguro que desea CERRAR la CUENTA del paciente que tiene un GARANTE Y DEUDA PENDIENTE", vbQuestion + vbYesNo, "Facturación") = vbYes Then
        Dim Login As New Login
        Login.UsuarioDeEstadoDeCuenta = ml_idUsuario
        Login.CargaDesdeOtraOpcion = True
        Login.Show vbModal
        If Not Login.Autenticado Or Login.IdUsuarioAutenticado <> ml_idUsuario Then
            Exit Sub
        End If
        mo_DOCuentaAtencion.IdUsuarioAuditoria = ml_idUsuario
        mo_DOCuentaAtencion.TotalPorPagar = CCur(txtTotalServicios.Text) + CCur(txtTotalFarmacia.Text)
        If mo_ReglasFacturacion.CuentasAtencionAltaConDeudaYGarante(mo_DOCuentaAtencion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtEstadoCuenta.Text) Then
            mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar ml_idCuentaAtencion, False, 0
            MsgBox "La Cuenta pasó a estado=CUENTA CERRADA, PENDIENTE DE PAGO CON GARANTE", vbInformation, "Facturación"
            LimparDatos
        Else
            MsgBox "No se pudo cerrar la cuenta", vbInformation, "Facturación"
        End If
    End If
End Sub

Private Sub btnCtaPagada_Click()
    If MsgBox("Esta seguro que la Cuenta pase a estado=PAGADA", vbQuestion + vbYesNo, "Facturación") = vbYes Then
        Dim Login As New Login
        Login.UsuarioDeEstadoDeCuenta = ml_idUsuario
        Login.CargaDesdeOtraOpcion = True
        Login.Show vbModal
        If Not Login.Autenticado Or Login.IdUsuarioAutenticado <> ml_idUsuario Then
            Exit Sub
        End If
        mo_DOCuentaAtencion.IdUsuarioAuditoria = ml_idUsuario
        If mo_ReglasFacturacion.CuentasAtencionPagada(mo_DOCuentaAtencion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtEstadoCuenta.Text) = True Then
            mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar ml_idCuentaAtencion, False, 0
            MsgBox "La Cuenta pasó a estado=PAGADA", vbInformation, "Facturación"
            LimparDatos
        Else
            MsgBox "No se pudo", vbInformation, "Facturación"
        End If
    End If
End Sub

'debb-21/10/2015
Sub EstadoCuentaConsolidadaXitem()
    If txtPaciente.Text = "" Then
        MsgBox "Tiene que LEER el Estado de Cuenta del Paciente", vbInformation, "Resultado"
    Else
        Dim oGenerarRecordsetProductos As New SighFacturacion.dllFactUcEstadoCuenta
        Dim rsReporte As New Recordset
        Dim rsItems As New Recordset
        Dim oConexion As New Connection
        Dim lnFor As Integer, ldFechaAlta As Date, lcHoraAlta As String, lcPuntoCarga As String
        Dim lnTotalPagarEstancia As Double, lnTotalDiasEstancia As Long, lnForNum As Integer
        Dim lcCodigo As String, lcItem As String, lnCantidad As Long, lnPrecio As Double, lnTotal As Double, lnTotalGen As Double
        Dim lbEsOpenOffice As Boolean, iFila As Long
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        With rsItems
              .Fields.Append "PuntoCarga", adVarChar, 50, adFldIsNullable
              .Fields.Append "codigo", adVarChar, 20, adFldIsNullable
              .Fields.Append "item", adVarChar, 200, adFldIsNullable
              .Fields.Append "cantidad", adInteger
              .Fields.Append "precio", adDouble
              .Fields.Append "total", adDouble
              .LockType = adLockOptimistic
              .Open
        End With
        'Farmacia
        Set rsReporte = ucFacturacionBienes.FacturacionProductos
        If rsReporte.RecordCount > 0 Then
           rsReporte.Sort = "IdPuntoCarga,codigo"
           rsReporte.MoveFirst
           Do While Not rsReporte.EOF
              lcCodigo = rsReporte.Fields("codigo").Value
              lcItem = rsReporte.Fields("NombreProducto").Value
              lnCantidad = 0: lnPrecio = 0: lnTotal = 0
              Select Case lnIdTipoFinanciamientoActual
              Case sghTipoFinanciamiento.sghPacienteNormal
                   lnPrecio = rsReporte!PrecioUnitario
              Case sghTipoFinanciamiento.sghSis
                   lnPrecio = rsReporte!precioSIS
              Case sghTipoFinanciamiento.sghSOAT
                   lnPrecio = rsReporte!PrecioSOAT
              Case sghTipoFinanciamiento.sghConvenios
                   lnPrecio = rsReporte!precioCONV
              End Select
              Do While Not rsReporte.EOF And lcCodigo = rsReporte.Fields("codigo").Value
                 lnCantidad = lnCantidad + rsReporte!CantidadPagar
                 rsReporte.MoveNext
                 If rsReporte.EOF Then
                    Exit Do
                 End If
              Loop
              rsItems.AddNew
              rsItems.Fields!PuntoCarga = "Farmacia"
              rsItems.Fields!Codigo = lcCodigo
              rsItems.Fields!Item = Left(lcItem, 200)
              rsItems.Fields!cantidad = lnCantidad
              rsItems.Fields!Precio = lnPrecio
              rsItems.Fields!Total = Round(lnCantidad * lnPrecio, 2)
              rsItems.Update
              
           Loop
        End If
        'Servicio
        lnFor = 1
        If ml_IdTipoServicio = sghHospitalizacion Then
           lnFor = 2
        End If
        For lnForNum = 1 To lnFor
            If lnForNum = 1 Then
                Set rsReporte = ucFacturacionServicios.FacturacionProductos
                If ml_IdTipoServicio = sghHospitalizacion Then
                   On Error Resume Next
                   rsReporte.Filter = "idPuntoCarga<>" & sghPtoCargaAdmisionHospitalizacion
                End If
            Else
                rsReporte.Filter = ""
                If txtFegreso.Text = "" Then
                    ldFechaAlta = CDate(Format(Now, sighentidades.DevuelveFechaSoloFormato_DMY))
                    lcHoraAlta = Format(Now, sighentidades.DevuelveHoraSoloFormato_HM)
                Else
                    ldFechaAlta = CDate(Format(CDate(txtFegreso.Text), sighentidades.DevuelveFechaSoloFormato_DMY))
                    lcHoraAlta = Format(CDate(txtFegreso.Text), sighentidades.DevuelveHoraSoloFormato_HM)
                End If
                oGenerarRecordsetProductos.GenerarRecordsetProductos rsReporte
                mo_AdminAdmision.GeneraEstanciaPorCadaServicioHospitalizado ml_idCuentaAtencion, ldFechaAlta, lcHoraAlta, rsReporte, lnTotalPagarEstancia, lnTotalDiasEstancia, oConexion
            End If
            '
            If rsReporte.RecordCount > 0 Then
                rsReporte.Sort = "IdPuntoCarga,codigo"
                rsReporte.MoveFirst
                Do While Not rsReporte.EOF
                   idPuntoCarga = rsReporte.Fields("IdPuntoCarga").Value
                   lcPuntoCarga = mo_ReglasComunes.FactPuntosCargaSeleccionarPorIdDevDescripcion(rsReporte.Fields("IdPuntoCarga").Value)
                   lcCodigo = rsReporte.Fields("codigo").Value
                   lcItem = rsReporte.Fields("NombreProducto").Value
                   lnCantidad = 0: lnPrecio = 0
                   Select Case lnIdTipoFinanciamientoActual
                   Case sghTipoFinanciamiento.sghPacienteNormal
                        lnPrecio = rsReporte!PrecioUnitario
                   Case sghTipoFinanciamiento.sghSis
                        lnPrecio = rsReporte!precioSIS
                   Case sghTipoFinanciamiento.sghSOAT
                        lnPrecio = rsReporte!PrecioSOAT
                   Case sghTipoFinanciamiento.sghConvenios
                        lnPrecio = rsReporte!precioCONV
                   End Select
                   Do While Not rsReporte.EOF And lcCodigo = rsReporte.Fields("codigo").Value And idPuntoCarga = rsReporte.Fields("IdPuntoCarga").Value
                      lnCantidad = lnCantidad + rsReporte!CantidadPagar
                      rsReporte.MoveNext
                      If rsReporte.EOF Then
                         Exit Do
                      End If
                   Loop
                   If sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica1 Or sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica2 Or _
                      sghPuntosCargaBasicos.sghPtoCargaBancoSangre1 Or sghPuntosCargaBasicos.sghPtoCargaBancoSangre2 Or _
                      sghPuntosCargaBasicos.sghPtoCargaEcogGeneral Or sghPuntosCargaBasicos.sghPtoCargaEcogObstetrica Or _
                      sghPuntosCargaBasicos.sghPtoCargaPatologiaClinica Or sghPuntosCargaBasicos.sghPtoCargaRayosX Or _
                      sghPuntosCargaBasicos.sghPtoCargaAdmisionHospitalizacion Then
                            rsItems.AddNew
                            rsItems.Fields!PuntoCarga = lcPuntoCarga
                            rsItems.Fields!Codigo = lcCodigo
                            rsItems.Fields!Item = Left(lcItem, 200)
                            rsItems.Fields!cantidad = lnCantidad
                            rsItems.Fields!Precio = lnPrecio
                            rsItems.Fields!Total = Round(lnCantidad * lnPrecio, 2)
                            rsItems.Update
                   End If
                Loop
             End If
        Next
        If rsItems.RecordCount > 0 Then
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
                Dim lnHWnd As Long
                'Abre el archivo ExcelOpenOffice
                lcArchivoExcel = App.Path + "\Plantillas\HojaLibre.ods"
        '        FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
        '        Chemin = "file:///" & App.Path & "\Plantillas\"
        '        Chemin = Replace(Chemin, "\", "/")
        '        Fichier = Chemin & "/OpenOffice.ods"
                '
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
                'Encabezado de Pagina
                mo_CabeceraReportes.CabeceraReportes Document, True
                ' Pone la ventana en primer plano, pasándole el Hwnd
                ret = SetForegroundWindow(lnHWnd)
                
            Else
                Dim oExcel As Excel.Application
                Dim oWorkBookPlantilla As Workbook
                Dim oWorkBook As Workbook
                Dim oWorkSheet As Worksheet
                Dim oRange As range
                Dim range As Excel.range
                Dim borders As Excel.borders
                'Crea nueva hoja
                Set oExcel = GalenhosExcelApplication()  'New Excel.Application
                Set oWorkBook = oExcel.Workbooks.Add
                'Abre, copia y cierra la plantilla
                Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HojaLibre.xls")
                oWorkBookPlantilla.Worksheets("Hoja_libre").Copy Before:=oWorkBook.Sheets(1)
                oWorkBookPlantilla.Close
                'Activa la primera hoja
                Set oWorkSheet = oWorkBook.Sheets(1)
                mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
                
           End If
        
           rsItems.Sort = "puntoCarga,item"
           
           
           If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(0, 0).setFormula("CONSUMOS CONSOLIDADOS de la Cuenta Nª: " & Trim(Str(ml_idCuentaAtencion)))
                Call Feuille.getcellbyposition(0, 1).setFormula("Estado de Cuenta al " & lcBuscaParametro.RetornaFechaServidorSQL())
                Call Feuille.getcellbyposition(0, 2).setFormula("Paciente: " & txtNroHistoria.Text & " - " & txtPaciente.Text)
                Call Feuille.getcellbyposition(0, 3).setFormula("Cuenta: " & txtEstadoCuenta.Text)
                Call Feuille.getcellbyposition(0, 4).setFormula("Serv.Egreso: " & txtServicio.Text)
           Else
                oWorkSheet.Cells(1, 1).Value = "CONSUMOS CONSOLIDADOS de la Cuenta Nª: " & Trim(Str(ml_idCuentaAtencion))
                oWorkSheet.Cells(1, 1).Font.Bold = True
                oWorkSheet.Cells(2, 1).Value = "Estado de Cuenta al " & lcBuscaParametro.RetornaFechaServidorSQL()
                oWorkSheet.Cells(3, 1).Value = "Paciente: " & txtNroHistoria.Text & " - " & txtPaciente.Text
                oWorkSheet.Cells(4, 1).Value = "Cuenta: " & txtEstadoCuenta.Text
                oWorkSheet.Cells(5, 1).Value = "Serv.Egreso: " & txtServicio.Text
           End If
           iFila = 7
           
           If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(0, iFila - 1).setFormula("PuntoCarga")
                Call Feuille.getcellbyposition(2, iFila - 1).setFormula("Código")
                Call Feuille.getcellbyposition(3, iFila - 1).setFormula("Descripción")
                Call Feuille.getcellbyposition(7, iFila - 1).setFormula("Cantidad")
                Call Feuille.getcellbyposition(8, iFila - 1).setFormula("Precio")
                Call Feuille.getcellbyposition(9, iFila - 1).setFormula("Total")
                Call Feuille.getcellbyposition(10, iFila - 1).setFormula("Tot.PCarga")
                Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(0) & CStr(iFila) & ":" & mo_ReglasReportes.BuscaNombreColumna(10) & CStr(iFila))
                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
           Else
                oWorkSheet.Cells(iFila, 1).Value = "PuntoCarga"
                oWorkSheet.Cells(iFila, 3).Value = "Código"
                oWorkSheet.Cells(iFila, 4).Value = "Descripción"
                oWorkSheet.Cells(iFila, 8).Value = "Cantidad"
                oWorkSheet.Cells(iFila, 9).Value = "Precio"
                oWorkSheet.Cells(iFila, 10).Value = "Total"
                oWorkSheet.Cells(iFila, 11).Value = "Tot.PCarga"
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 1, iFila, 11
           End If
           iFila = iFila + 1
           
           rsItems.MoveFirst
           lnTotalGen = 0
           Do While Not rsItems.EOF
              lcPuntoCarga = rsItems!PuntoCarga
              lnTotal = 0
              If lbEsOpenOffice = True Then
                     Call Feuille.getcellbyposition(0, iFila - 1).setFormula(rsItems!PuntoCarga)
              Else
                     oWorkSheet.Cells(iFila, 1).Value = rsItems!PuntoCarga
              End If
              Do While Not rsItems.EOF And lcPuntoCarga = rsItems!PuntoCarga
                    If lbEsOpenOffice = True Then
                           Call Feuille.getcellbyposition(2, iFila - 1).setFormula(rsItems!Codigo)
                           Call Feuille.getcellbyposition(3, iFila - 1).setFormula(rsItems!Item)
                           Call Feuille.getcellbyposition(7, iFila - 1).setFormula(rsItems!cantidad)
                           Call Feuille.getcellbyposition(8, iFila - 1).setFormula(rsItems!Precio)
                           Call Feuille.getcellbyposition(9, iFila - 1).setFormula(rsItems!Total)
                    Else
                           oWorkSheet.Cells(iFila, 3).Value = rsItems!Codigo
                           oWorkSheet.Cells(iFila, 4).Value = rsItems!Item
                           oWorkSheet.Cells(iFila, 4).Font.Name = "Arial Narrow"
                           oWorkSheet.Cells(iFila, 4).Font.Size = 8
                           oWorkSheet.Cells(iFila, 8).Value = rsItems!cantidad
                           oWorkSheet.Cells(iFila, 9).Value = rsItems!Precio
                           oWorkSheet.Cells(iFila, 10).Value = rsItems!Total
                    End If
                    iFila = iFila + 1
                    lnTotal = lnTotal + rsItems!Total
                    lnTotalGen = lnTotalGen + rsItems!Total
                    rsItems.MoveNext
                    If rsItems.EOF Then
                       Exit Do
                    End If
              Loop
              iFila = iFila - 1
              If lbEsOpenOffice = True Then
                     Call Feuille.getcellbyposition(10, iFila - 1).setFormula(lnTotal)
              Else
                     oWorkSheet.Cells(iFila, 11).Value = lnTotal
              End If
              iFila = iFila + 2
           Loop
           iFila = iFila + 1
           If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(10, iFila - 1).setFormula(lnTotalGen)
                Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(0) & CStr(iFila) & ":" & mo_ReglasReportes.BuscaNombreColumna(10) & CStr(iFila))
                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
           Else
                oWorkSheet.Cells(iFila, 11).Value = lnTotalGen
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 1, iFila, 11
           End If
           If lbEsOpenOffice = True Then
                'MsgBox "Se generó en forma correcta el reporte: " & lcArchivoExcel, vbInformation
                Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                PrintArea(0).Sheet = 0
                PrintArea(0).startcolumn = 1
                PrintArea(0).StartRow = 0
                PrintArea(0).EndColumn = 17
                PrintArea(0).EndRow = iFila
    '            Call Feuille.SetPrintAreas(PrintArea())
    '            Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
                Call Feuille.SetPrintAreas(PrintArea())
                Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
                MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
            Else
                oWorkSheet.PageSetup.Orientation = xlLandscape
                oWorkSheet.PageSetup.PrintTitleRows = "$1:$7"
                If oWorkSheet.PageSetup.PrintArea <> "" Then oWorkSheet.PageSetup.PrintArea = "$A$1:$R$" & (iFila + 2) 'sighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
                oExcel.Visible = True
                oWorkSheet.PrintPreview
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
                'encabezado de pagina
                Set PageStyles = Nothing
                Set Sheet = Nothing
                Set StyleFamilies = Nothing
                Set DefPage = Nothing
                Set Htext = Nothing
                Set Hcontent = Nothing
           Else
                'liberar memoria
                Set oExcel = Nothing
                Set oWorkBookPlantilla = Nothing
                Set oWorkBook = Nothing
                Set oWorkSheet = Nothing
           End If
           
           
        End If
    End If

    Set oGenerarRecordsetProductos = Nothing
    Set rsReporte = Nothing
    Set rsItems = Nothing
    Set oConexion = Nothing
End Sub

'***************daniel barrantes**************
'***************Impresion del Estado de Cuenta
'***************
Private Sub btnImprimir_Click()
    'debb-21/10/2015 (inicio)
    If MsgBox("CONSUMOS CONSOLIDADOS ?", vbQuestion + vbYesNo, "") = vbYes Then
       EstadoCuentaConsolidadaXitem
       Exit Sub
    End If
    'debb-21/10/2015 (fin)
    
    Dim iFila As Long: Dim iCol As Integer
    Dim ms_EstadosFacturacion As String
    Dim ms_TiposFinanciamiento As String
    Dim ml_AgruparPor As Long
    Dim mo_ReporteUtil As New sighentidades.ReporteUtil
    Dim idPuntoCarga As Long
    
    Dim lnTSubTotal As Double: Dim lnTSubTotalSIS As Double: Dim lnTSubTotalSOAT As Double: Dim lnTSubTotalEXO As Double: Dim lnTsubTotalConvenio As Double
    Dim lnTSubTotalPAGO As Double: Dim lnTSubTotalDEBE As Double: Dim lnTSubTotalSALDO As Double
    
    Dim lnTotal As Double: Dim lnTotalSIS As Double: Dim lnTotalSOAT As Double: Dim lnTotalEXO As Double: Dim lnTotalConvenio As Double
    Dim lnTotalPAGO As Double: Dim lnTotalDEBE As Double: Dim lnTotalSALDO As Double
    
    Dim lnDebe As Double: Dim lnPago As Double: Dim lnSaldo As Double
    Dim lnTDebe As Double: Dim lnTPago As Double: Dim lnTSaldo As Double
    Dim lnSIS As Double: Dim lnSOAT As Double: Dim lnEXO As Double: Dim lnTotalCredito As Double: Dim lnConvenio As Double
    Dim lnDctos As Double: Dim lnPagoEnFarmacia As Double: Dim lnPagoEnServicios As Double
    Dim lnTotalPagosAdelantados As Double
    Dim lnCantidadPagarBienes As Long, lnTotalPagarBienes As Double
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim ldFechaAlta As Date, lcHoraAlta As String
    Dim lnFor As Integer, lnForNum As Integer
    Dim oGenerarRecordsetProductos As New SighFacturacion.dllFactUcEstadoCuenta
    Dim lnTotalPagarEstancia As Double, lnTotalDiasEstancia As Long
    Dim lbEsOpenOffice As Boolean
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    
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
        Dim lnHWnd As Long
    Else
        Dim oExcel As Excel.Application
        Dim oWorkBookPlantilla As Workbook
        Dim oWorkBook As Workbook
        Dim oWorkSheet As Worksheet
        Dim oRange As range
        Dim range As Excel.range
        Dim borders As Excel.borders
    End If
        
    If txtPaciente.Text = "" Then
        MsgBox "Tiene que LEER el Estado de Cuenta del Paciente", vbInformation, "Resultado"
    Else
        If ucFacturacionBienes.FacturacionProductos.RecordCount = 0 And ucFacturacionServicios.FacturacionProductos.RecordCount = 0 Then
           MsgBox "Tiene que LEER el Estado de Cuenta del Paciente", vbInformation, "Resultado"
           Exit Sub
        End If
        MousePointer = 11
        
        If lbEsOpenOffice = True Then
            'Abre el archivo ExcelOpenOffice
            lcArchivoExcel = App.Path + "\Plantillas\ECuentaCte.ods"
    '        FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
    '        Chemin = "file:///" & App.Path & "\Plantillas\"
    '        Chemin = Replace(Chemin, "\", "/")
    '        Fichier = Chemin & "/OpenOffice.ods"
            '
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
            'Encabezado de Pagina
            mo_CabeceraReportes.CabeceraReportes Document, True
            ' Pone la ventana en primer plano, pasándole el Hwnd
            ret = SetForegroundWindow(lnHWnd)
        Else
            'Crea nueva hoja
            Set oExcel = GalenhosExcelApplication()  'New Excel.Application
            Set oWorkBook = oExcel.Workbooks.Add
            'Abre, copia y cierra la plantilla
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\ECuentaCte.xls")
            oWorkBookPlantilla.Worksheets("CuentaCte").Copy Before:=oWorkBook.Sheets(1)
            oWorkBookPlantilla.Close
            'Activa la primera hoja
            Set oWorkSheet = oWorkBook.Sheets(1)
            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
         End If
        
        'Inicio de Impresion
        
        If ml_IdTipoServicio = sghConsultaExterna Then
           Set rsReporte = mo_ReglasFacturacion.AtencionesFiltraIngresosPorIdAtencion(ml_idAtencion)
        Else
           Set rsReporte = mo_ReglasFacturacion.AtencionesFiltraEgresosPorIdAtencion(ml_idAtencion)
        End If
        If rsReporte.RecordCount = 0 Then
           MousePointer = 1
           Exit Sub
        End If
        
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(2, 0).setFormula("Estado de Cuenta al " & lcBuscaParametro.RetornaFechaServidorSQL())
            Call Feuille.getcellbyposition(2, 2).setFormula(Trim(Str(ml_idCuentaAtencion)) & "  " & txtEstadoCuenta.Text)
            Call Feuille.getcellbyposition(7, 2).setFormula(Trim(Str(ml_idAtencion)) & IIf(txtDxEgr.Text = "", "", "      Dx Egreso: " & txtDxEgr.Text))
            Call Feuille.getcellbyposition(2, 3).setFormula(txtPaciente.Text)
            Call Feuille.getcellbyposition(7, 3).setFormula(Trim(txtNroHistoria.Text) & "       Dom.Pac: " & Trim(txtDomicilioPacienteEnAtencion.Text))
            Call Feuille.getcellbyposition(2, 4).setFormula("'" & Format(rsReporte.Fields!FechaIngreso, sighentidades.DevuelveFechaSoloFormato_DMY) & " " & rsReporte.Fields!HoraIngreso)
            Call Feuille.getcellbyposition(7, 4).setFormula(IIf(IsNull(rsReporte.Fields!codServicio), "", rsReporte.Fields!codServicio & " - " & rsReporte.Fields!DServicio))
            Call Feuille.getcellbyposition(2, 5).setFormula(IIf(IsNull(rsReporte.Fields!FechaEgreso), "", "'" & Format(rsReporte.Fields!FechaEgreso & " " & rsReporte.Fields!HoraEgreso, sighentidades.DevuelveFechaSoloFormato_DMY_HM)))
            Call Feuille.getcellbyposition(7, 5).setFormula(IIf(IsNull(rsReporte.Fields!codCama), "", rsReporte.Fields!codCama))
        Else
            oWorkSheet.Cells(1, 3).Value = "Estado de Cuenta al " & lcBuscaParametro.RetornaFechaServidorSQL()
            oWorkSheet.Cells(3, 3).Value = Trim(Str(ml_idCuentaAtencion)) & "  " & txtEstadoCuenta.Text
            oWorkSheet.Cells(3, 8).Value = Trim(Str(ml_idAtencion)) & IIf(txtDxEgr.Text = "", "", "      Dx Egreso: " & txtDxEgr.Text)
            oWorkSheet.Cells(4, 3).Value = txtPaciente.Text
            oWorkSheet.Cells(4, 8).Value = Trim(txtNroHistoria.Text) & "       Dom.Pac: " & Trim(txtDomicilioPacienteEnAtencion.Text)
            oWorkSheet.Cells(5, 3).Value = "'" & Format(rsReporte.Fields!FechaIngreso, sighentidades.DevuelveFechaSoloFormato_DMY) & " " & rsReporte.Fields!HoraIngreso
            oWorkSheet.Cells(5, 8).Value = IIf(IsNull(rsReporte.Fields!codServicio), "", rsReporte.Fields!codServicio & " - " & rsReporte.Fields!DServicio)
            oWorkSheet.Cells(6, 3).Value = IIf(IsNull(rsReporte.Fields!FechaEgreso), "", "'" & Format(rsReporte.Fields!FechaEgreso & " " & rsReporte.Fields!HoraEgreso, sighentidades.DevuelveFechaSoloFormato_DMY_HM))
            oWorkSheet.Cells(6, 8).Value = IIf(IsNull(rsReporte.Fields!codCama), "", rsReporte.Fields!codCama)
        End If
        
        iFila = 8
        iCol = 2
        ms_EstadosFacturacion = ""
        ms_TiposFinanciamiento = ""
        ml_AgruparPor = 1
        lnTotal = 0: lnTotalSIS = 0: lnTotalSOAT = 0: lnTotalEXO = 0: lnTotalPAGO = 0: lnTotalDEBE = 0: lnTotalSALDO = 0: lnTotalCredito = 0
        lnTDebe = 0: lnTPago = 0: lnTSaldo = 0
        'Farmacia
        Set rsReporte = ucFacturacionBienes.FacturacionProductos
        If rsReporte.RecordCount > 0 Then
            rsReporte.MoveFirst
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Farmacia")
            Else
                oWorkSheet.Cells(iFila, 2).Value = "Farmacia"
            End If
            lnTSubTotal = 0: lnTSubTotalSIS = 0: lnTSubTotalSOAT = 0: lnTSubTotalEXO = 0: lnTsubTotalConvenio = 0
            lnTSubTotalPAGO = 0: lnTSubTotalDEBE = 0: lnTSubTotalSALDO = 0
            Do While Not rsReporte.EOF
                If rsReporte.Fields!IdOrden = 37 Or rsReporte.Fields!IdOrden = 38 Then
                lnSIS = 0
                End If
               If rsReporte.Fields!idestadofacturacion = 4 Or rsReporte.Fields!idestadofacturacion = 1 Then   'Solo PAGADOS y REGISTRADOS
                    lnSIS = IIf(IsNull(rsReporte.Fields!ImporteSIS), 0, rsReporte.Fields!ImporteSIS)
                    lnEXO = IIf(IsNull(rsReporte.Fields!importeEXO), 0, rsReporte.Fields!importeEXO)
                    lnSOAT = IIf(IsNull(rsReporte.Fields!ImporteSOAT), 0, rsReporte.Fields!ImporteSOAT)
                    lnConvenio = IIf(IsNull(rsReporte.Fields!ImporteConv), 0, rsReporte.Fields!ImporteConv)
                    lnCantidadPagarBienes = rsReporte.Fields("CantidadPagar").Value  'rsReporte.Fields("CantidadPagar").Value - rsReporte.Fields!CantidadDevuelta
                    lnTotalPagarBienes = rsReporte.Fields("TotalPagar").Value  'Round(lnCantidadPagarBienes * rsReporte.Fields("preciounitario").Value, 2)
                    
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(2, iFila - 1).setFormula(rsReporte.Fields("Codigo").Value & " - " & rsReporte.Fields("NombreProducto").Value)
                        Call Feuille.getcellbyposition(4, iFila - 1).setFormula(lnCantidadPagarBienes)
                        Call Feuille.getcellbyposition(5, iFila - 1).setFormula(rsReporte.Fields("preciounitario").Value)
                        Call Feuille.getcellbyposition(6, iFila - 1).setFormula(Format(lnTotalPagarBienes, "####,###.#0"))
                        Call Feuille.getcellbyposition(7, iFila - 1).setFormula(Format(lnEXO, "####,###.#0"))
                        Call Feuille.getcellbyposition(8, iFila - 1).setFormula(Format(lnSIS, "####,###.#0"))
                        Call Feuille.getcellbyposition(9, iFila - 1).setFormula(Format(lnSOAT, "####,###.#0"))
                        Call Feuille.getcellbyposition(10, iFila - 1).setFormula(Format(lnConvenio, "####,###.#0"))
                    Else
                        oWorkSheet.Cells(iFila, 3).Value = rsReporte.Fields("Codigo").Value & " - " & rsReporte.Fields("NombreProducto").Value
                        oWorkSheet.Cells(iFila, 5).Value = lnCantidadPagarBienes
                        oWorkSheet.Cells(iFila, 6).Value = rsReporte.Fields("preciounitario").Value
                        oWorkSheet.Cells(iFila, 7).Value = Format(lnTotalPagarBienes, "####,###.#0")
                        oWorkSheet.Cells(iFila, 8).Value = Format(lnEXO, "####,###.#0")
                        oWorkSheet.Cells(iFila, 9).Value = Format(lnSIS, "####,###.#0")
                        oWorkSheet.Cells(iFila, 10).Value = Format(lnSOAT, "####,###.#0")
                        oWorkSheet.Cells(iFila, 11).Value = Format(lnConvenio, "####,###.#0")
                    End If
                    
                    If lbGeneraReciboPago = True Then
                       lnDebe = lnTotalPagarBienes - lnEXO - lnSIS - lnSOAT
                    Else
                       If (rsReporte.Fields!CantidadSIS + rsReporte.Fields!CantidadSOAT + rsReporte.Fields!cantidadConv) = 0 Then
                           lnDebe = lnTotalPagarBienes
                       Else
                           lnDebe = 0
                       End If
                    End If
                    If rsReporte.Fields!idestadofacturacion = 4 Then
                       lnPago = rsReporte.Fields("ImporteEnBoleta").Value - lnEXO
                    Else
                       lnPago = 0
                    End If

                    lnSaldo = lnDebe - lnPago
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnDebe, "####,###.#0"))
                        Call Feuille.getcellbyposition(12, iFila - 1).setFormula(Format(lnPago, "####,###.#0"))
                        Call Feuille.getcellbyposition(13, iFila - 1).setFormula(Format(lnSaldo, "####,###.#0"))
                        Call Feuille.getcellbyposition(14, iFila - 1).setFormula(rsReporte.Fields!nroDcto)
                        Call Feuille.getcellbyposition(15, iFila - 1).setFormula(rsReporte.Fields!FechaDespacho)
                        Call Feuille.getcellbyposition(16, iFila - 1).setFormula(rsReporte.Fields!ServicioDeEstancia)
                    Else
                        oWorkSheet.Cells(iFila, 12).Value = Format(lnDebe, "####,###.#0")
                        oWorkSheet.Cells(iFila, 13).Value = Format(lnPago, "####,###.#0")
                        oWorkSheet.Cells(iFila, 14).Value = Format(lnSaldo, "####,###.#0")
                        oWorkSheet.Cells(iFila, 15).Value = rsReporte.Fields!nroDcto    'movNumero
                        oWorkSheet.Cells(iFila, 16).Value = rsReporte.Fields!FechaDespacho
                        oWorkSheet.Cells(iFila, 17).Value = rsReporte.Fields!ServicioDeEstancia
                    End If
                    
                    lnTSubTotal = lnTSubTotal + lnTotalPagarBienes
                    lnTSubTotalEXO = lnTSubTotalEXO + lnEXO
                    lnTSubTotalSIS = lnTSubTotalSIS + lnSIS
                    lnTSubTotalSOAT = lnTSubTotalSOAT + lnSOAT
                    lnTsubTotalConvenio = lnTsubTotalConvenio + lnConvenio
                    lnTSubTotalPAGO = lnTSubTotalPAGO + lnPago
                    lnTSubTotalDEBE = lnTSubTotalDEBE + lnDebe
                    lnTSubTotalSALDO = lnTSubTotalSALDO + lnSaldo
                    
                    lnTotal = lnTotal + lnTotalPagarBienes
                    lnTotalSIS = lnTotalSIS + lnSIS
                    lnTotalSOAT = lnTotalSOAT + lnSOAT
                    lnTotalEXO = lnTotalEXO + lnEXO
                    lnTotalConvenio = lnTotalConvenio + lnConvenio
                    lnTotalPAGO = lnTotalPAGO + lnPago
                    lnTotalDEBE = lnTotalDEBE + lnDebe
                    lnTotalSALDO = lnTotalSALDO + lnSaldo
            
                    If rsReporte.Fields!idProducto = lnIdPagosACuenta Then   'Pagos a cuenta
                       lnTotalCredito = lnTotalCredito + lnTotalPagarBienes
                    End If
                    
                    iFila = iFila + 1
                End If
                rsReporte.MoveNext
             Loop
            If lbEsOpenOffice = True Then
                Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(3) & CStr(iFila) & ":" & mo_ReglasReportes.BuscaNombreColumna(iCol + 14) & CStr(iFila))
                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
            Else
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 3, iFila, iCol + 14
            End If
            
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(6, iFila - 1).setFormula(Format(lnTSubTotal, "####,###.#0"))
                Call Feuille.getcellbyposition(7, iFila - 1).setFormula(Format(lnTSubTotalEXO, "####,###.#0"))
                Call Feuille.getcellbyposition(8, iFila - 1).setFormula(Format(lnTSubTotalSIS, "####,###.#0"))
                Call Feuille.getcellbyposition(9, iFila - 1).setFormula(Format(lnTSubTotalSOAT, "####,###.#0"))
                Call Feuille.getcellbyposition(10, iFila - 1).setFormula(Format(lnTsubTotalConvenio, "####,###.#0"))
                Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnTSubTotalDEBE, "####,###.#0"))
                Call Feuille.getcellbyposition(12, iFila - 1).setFormula(Format(lnTSubTotalPAGO, "####,###.#0"))
                Call Feuille.getcellbyposition(13, iFila - 1).setFormula(Format(lnTSubTotalSALDO, "####,###.#0"))
            Else
                oWorkSheet.Cells(iFila, 7).Value = Format(lnTSubTotal, "####,###.#0")
                oWorkSheet.Cells(iFila, 8).Value = Format(lnTSubTotalEXO, "####,###.#0")
                oWorkSheet.Cells(iFila, 9).Value = Format(lnTSubTotalSIS, "####,###.#0")
                oWorkSheet.Cells(iFila, 10).Value = Format(lnTSubTotalSOAT, "####,###.#0")
                oWorkSheet.Cells(iFila, 11).Value = Format(lnTsubTotalConvenio, "####,###.#0")
                oWorkSheet.Cells(iFila, 12).Value = Format(lnTSubTotalDEBE, "####,###.#0")
                oWorkSheet.Cells(iFila, 13).Value = Format(lnTSubTotalPAGO, "####,###.#0")
                oWorkSheet.Cells(iFila, 14).Value = Format(lnTSubTotalSALDO, "####,###.#0")
            End If
                
            iFila = iFila + 1
        End If
        rsReporte.Close
        lnPagoEnFarmacia = lnTSubTotalSALDO
        'Servicios
        lnTotalPagarEstancia = 0
        lnFor = 1
        If ml_IdTipoServicio = sghHospitalizacion Then
           lnFor = 2
        End If
        For lnForNum = 1 To lnFor
            If lnForNum = 1 Then
                Set rsReporte = ucFacturacionServicios.FacturacionProductos
                If ml_IdTipoServicio = sghHospitalizacion Then
                   On Error Resume Next
                   rsReporte.Filter = "idPuntoCarga<>" & sghPtoCargaAdmisionHospitalizacion
                End If
            Else
                rsReporte.Filter = ""
                If txtFegreso.Text = "" Then
                    ldFechaAlta = CDate(Format(Now, sighentidades.DevuelveFechaSoloFormato_DMY))
                    lcHoraAlta = Format(Now, sighentidades.DevuelveHoraSoloFormato_HM)
                Else
                    ldFechaAlta = CDate(Format(CDate(txtFegreso.Text), sighentidades.DevuelveFechaSoloFormato_DMY))
                    lcHoraAlta = Format(CDate(txtFegreso.Text), sighentidades.DevuelveHoraSoloFormato_HM)
                End If
                oGenerarRecordsetProductos.GenerarRecordsetProductos rsReporte
                mo_AdminAdmision.GeneraEstanciaPorCadaServicioHospitalizado ml_idCuentaAtencion, ldFechaAlta, lcHoraAlta, rsReporte, lnTotalPagarEstancia, lnTotalDiasEstancia, oConexion
                If txtFegreso.Text <> "" Then
                   lnTotalPagarEstancia = 0
                End If
            End If
            '
            lnTotalPagosAdelantados = 0
            If rsReporte.RecordCount > 0 Then
                rsReporte.Sort = "IdPuntoCarga"
                rsReporte.MoveFirst
                Do While Not rsReporte.EOF
                    idPuntoCarga = rsReporte.Fields("IdPuntoCarga").Value
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(1, iFila - 1).setFormula(mo_ReglasComunes.FactPuntosCargaSeleccionarPorIdDevDescripcion(rsReporte.Fields("IdPuntoCarga").Value))
                    Else
                        oWorkSheet.Cells(iFila, 2).Value = mo_ReglasComunes.FactPuntosCargaSeleccionarPorIdDevDescripcion(rsReporte.Fields("IdPuntoCarga").Value)
                    End If
                    'oWorkSheet.Cells(iFila, 2).Value = FactPuntosCargaSeleccionarPorId(rsReporte.Fields("IdPuntoCarga").Value)
                    
                    lnTSubTotal = 0: lnTSubTotalSIS = 0: lnTSubTotalSOAT = 0: lnTSubTotalEXO = 0: lnTsubTotalConvenio = 0
                    lnTSubTotalPAGO = 0: lnTSubTotalDEBE = 0: lnTSubTotalSALDO = 0
                    Do While Not rsReporte.EOF And idPuntoCarga = rsReporte.Fields("IdPuntoCarga").Value
                        If rsReporte.Fields!IdOrden = 641 Then
                        lnSIS = 0
                        End If
                        If rsReporte.Fields!idestadofacturacion = 4 Or rsReporte.Fields!idestadofacturacion = 1 Or rsReporte.Fields!idestadofacturacion = 10 Or rsReporte.Fields!idestadofacturacion = sghConPreVenta Then 'Solo PAGADOS/REGISTRADOS/AUTORIZ.AUTOMATICA/preventa                            lnSIS = IIf(IsNull(rsReporte.Fields!ImporteSIS), 0, rsReporte.Fields!ImporteSIS)
                            lnSIS = IIf(IsNull(rsReporte.Fields!ImporteSIS), 0, rsReporte.Fields!ImporteSIS)
                            lnSOAT = IIf(IsNull(rsReporte.Fields!ImporteSOAT), 0, rsReporte.Fields!ImporteSOAT)
                            lnEXO = IIf(IsNull(rsReporte.Fields!importeEXO), 0, rsReporte.Fields!importeEXO)
                            lnConvenio = IIf(IsNull(rsReporte.Fields!ImporteConv), 0, rsReporte.Fields!ImporteConv)
                            
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(2, iFila - 1).setFormula(rsReporte.Fields("Codigo").Value & " - " & rsReporte.Fields("NombreProducto").Value)
                                Call Feuille.getcellbyposition(4, iFila - 1).setFormula(rsReporte.Fields("CantidadPagar").Value)
                                Call Feuille.getcellbyposition(5, iFila - 1).setFormula(Format(rsReporte.Fields("preciounitario").Value, "####,###.###0"))
                                Call Feuille.getcellbyposition(6, iFila - 1).setFormula(Format(rsReporte.Fields("TotalPagar").Value, "####,###.#0"))
                                Call Feuille.getcellbyposition(7, iFila - 1).setFormula(Format(lnEXO, "####,###.#0"))
                                Call Feuille.getcellbyposition(8, iFila - 1).setFormula(Format(lnSIS, "####,###.#0"))
                                Call Feuille.getcellbyposition(9, iFila - 1).setFormula(Format(lnSOAT, "####,###.#0"))
                                Call Feuille.getcellbyposition(10, iFila - 1).setFormula(Format(lnConvenio, "####,###.#0"))
                            Else
                                oWorkSheet.Cells(iFila, 3).Value = rsReporte.Fields("Codigo").Value & " - " & rsReporte.Fields("NombreProducto").Value
                                oWorkSheet.Cells(iFila, 5).Value = rsReporte.Fields("CantidadPagar").Value
                                oWorkSheet.Cells(iFila, 6).Value = Format(rsReporte.Fields("preciounitario").Value, "####,###.###0")
                                oWorkSheet.Cells(iFila, 7).Value = Format(rsReporte.Fields("TotalPagar").Value, "####,###.#0")
                                oWorkSheet.Cells(iFila, 8).Value = Format(lnEXO, "####,###.#0")
                                oWorkSheet.Cells(iFila, 9).Value = Format(lnSIS, "####,###.#0")
                                oWorkSheet.Cells(iFila, 10).Value = Format(lnSOAT, "####,###.#0")
                                oWorkSheet.Cells(iFila, 11).Value = Format(lnConvenio, "####,###.#0")
                            End If
                            
                            If rsReporte.Fields!idestadofacturacion = 4 Then
                                lnPago = rsReporte.Fields("ImporteEnBoleta").Value - lnEXO
                            Else
                                lnPago = 0
                            End If
                            If rsReporte.Fields!idProducto <> lnIdPagosACuenta And rsReporte.Fields!idProducto <> lnIdPagosXdevoluciones Then
                               If lbGeneraReciboPago = True Then
                                    lnDebe = rsReporte.Fields!TotalPagar - lnEXO - lnSIS - lnSOAT
                               Else
                                    If (rsReporte.Fields!CantidadSIS + rsReporte.Fields!CantidadSOAT + rsReporte.Fields!cantidadConv) = 0 Then
                                        lnDebe = rsReporte.Fields!TotalPagar
                                    Else
                                        lnDebe = 0
                                    End If
                               End If
                               lnSaldo = lnDebe - lnPago
                            Else
                               If rsReporte.Fields!idProducto = lnIdPagosACuenta Then
                                    lnTotalPagosAdelantados = lnTotalPagosAdelantados + rsReporte.Fields!ImporteEnBoleta
                                    lnDebe = 0
                                    If lbGeneraReciboPago = True Then
                                       lnSaldo = -rsReporte.Fields!ImporteEnBoleta
                                    Else
                                       lnSaldo = 0
                                    End If
                               Else
                                    lnDebe = 0
                                    If lbGeneraReciboPago = True Then
                                       lnSaldo = rsReporte.Fields!ImporteEnBoleta
                                       lnPago = -rsReporte.Fields!ImporteEnBoleta
                                    Else
                                       lnSaldo = 0
                                       lnPago = 0
                                    End If
                               End If
                            End If
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnDebe, "####,###.#0"))
                                Call Feuille.getcellbyposition(12, iFila - 1).setFormula(Format(lnPago, "####,###.#0"))
                                Call Feuille.getcellbyposition(13, iFila - 1).setFormula(Format(lnSaldo, "####,###.#0"))
                                Call Feuille.getcellbyposition(14, iFila - 1).setFormula(rsReporte.Fields!nroDcto)
                                Call Feuille.getcellbyposition(15, iFila - 1).setFormula(rsReporte.Fields!FechaDespacho)
                                Call Feuille.getcellbyposition(16, iFila - 1).setFormula(rsReporte.Fields!ServicioDeEstancia)
                            Else
                                oWorkSheet.Cells(iFila, 12).Value = Format(lnDebe, "####,###.#0")
                                oWorkSheet.Cells(iFila, 13).Value = Format(lnPago, "####,###.#0")
                                oWorkSheet.Cells(iFila, 14).Value = Format(lnSaldo, "####,###.#0")
                                oWorkSheet.Cells(iFila, 15).Value = rsReporte.Fields!nroDcto
                                oWorkSheet.Cells(iFila, 16).Value = rsReporte.Fields!FechaDespacho
                                oWorkSheet.Cells(iFila, 17).Value = rsReporte.Fields!ServicioDeEstancia
                            End If
    
                            If rsReporte.Fields!idProducto <> lnIdPagosACuenta Then
                               lnTSubTotal = lnTSubTotal + rsReporte.Fields("TotalPagar").Value
                               lnTSubTotalDEBE = lnTSubTotalDEBE + lnDebe
                            End If
                            lnTSubTotalEXO = lnTSubTotalEXO + lnEXO
                            lnTSubTotalSIS = lnTSubTotalSIS + lnSIS
                            lnTSubTotalSOAT = lnTSubTotalSOAT + lnSOAT
                            lnTsubTotalConvenio = lnTsubTotalConvenio + lnConvenio
                            lnTSubTotalPAGO = lnTSubTotalPAGO + lnPago
                            lnTSubTotalSALDO = lnTSubTotalSALDO + lnSaldo
                            
                            If rsReporte.Fields!idProducto <> lnIdPagosACuenta Then
                               lnTotal = lnTotal + rsReporte.Fields("TotalPagar").Value
                               lnTotalDEBE = lnTotalDEBE + lnDebe
                            End If
                            lnTotalSIS = lnTotalSIS + lnSIS
                            lnTotalSOAT = lnTotalSOAT + lnSOAT
                            lnTotalEXO = lnTotalEXO + lnEXO
                            lnTotalConvenio = lnTotalConvenio + lnConvenio
                            lnTotalPAGO = lnTotalPAGO + lnPago
                            lnTotalSALDO = lnTotalSALDO + lnSaldo
                            
                            If rsReporte.Fields!idProducto = lnIdPagosACuenta Then   'Pagos a cuenta
                               lnTotalCredito = lnTotalCredito + rsReporte.Fields("TotalPorPagar").Value
                            End If
                            
                            iFila = iFila + 1
                        End If
                        rsReporte.MoveNext
                        If rsReporte.EOF Then
                           Exit Do
                        End If
                    Loop
                    If lbEsOpenOffice = True Then
                        Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(3) & CStr(iFila) & ":" & mo_ReglasReportes.BuscaNombreColumna(iCol + 14) & CStr(iFila))
                        mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                    Else
                        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 3, iFila, iCol + 14
                    End If
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(6, iFila - 1).setFormula(Format(lnTSubTotal, "####,###.#0"))
                        Call Feuille.getcellbyposition(7, iFila - 1).setFormula(Format(lnTSubTotalEXO, "####,###.#0"))
                        Call Feuille.getcellbyposition(8, iFila - 1).setFormula(Format(lnTSubTotalSIS, "####,###.#0"))
                        Call Feuille.getcellbyposition(9, iFila - 1).setFormula(Format(lnTSubTotalSOAT, "####,###.#0"))
                        Call Feuille.getcellbyposition(10, iFila - 1).setFormula(Format(lnTsubTotalConvenio, "####,###.#0"))
                        Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnTSubTotalDEBE, "####,###.#0"))
                        Call Feuille.getcellbyposition(12, iFila - 1).setFormula(Format(lnTSubTotalPAGO, "####,###.#0"))
                        Call Feuille.getcellbyposition(13, iFila - 1).setFormula(Format(lnTSubTotalSALDO, "####,###.#0"))
                    Else
                        oWorkSheet.Cells(iFila, 7).Value = Format(lnTSubTotal, "####,###.#0")
                        oWorkSheet.Cells(iFila, 8).Value = Format(lnTSubTotalEXO, "####,###.#0")
                        oWorkSheet.Cells(iFila, 9).Value = Format(lnTSubTotalSIS, "####,###.#0")
                        oWorkSheet.Cells(iFila, 10).Value = Format(lnTSubTotalSOAT, "####,###.#0")
                        oWorkSheet.Cells(iFila, 11).Value = Format(lnTsubTotalConvenio, "####,###.#0")
                        oWorkSheet.Cells(iFila, 12).Value = Format(lnTSubTotalDEBE, "####,###.#0")
                        oWorkSheet.Cells(iFila, 13).Value = Format(lnTSubTotalPAGO, "####,###.#0")
                        oWorkSheet.Cells(iFila, 14).Value = Format(lnTSubTotalSALDO, "####,###.#0")
                    End If
                    iFila = iFila + 1
                 Loop
            End If
            iFila = iFila + 1
            If lbEsOpenOffice = True Then
                Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(2) & CStr(iFila) & ":" & mo_ReglasReportes.BuscaNombreColumna(iCol + 14) & CStr(iFila))
                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
            Else
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, iCol + 14
            End If
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Total: ")
                Call Feuille.getcellbyposition(6, iFila - 1).setFormula(Format(lnTotal, "####,###.#0"))
                Call Feuille.getcellbyposition(7, iFila - 1).setFormula(Format(lnTotalEXO, "####,###.#0"))
                Call Feuille.getcellbyposition(8, iFila - 1).setFormula(Format(lnTotalSIS, "####,###.#0"))
                Call Feuille.getcellbyposition(9, iFila - 1).setFormula(Format(lnTotalSOAT, "####,###.#0"))
                Call Feuille.getcellbyposition(10, iFila - 1).setFormula(Format(lnTotalConvenio, "####,###.#0"))
                Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnTotalDEBE, "####,###.#0"))
                Call Feuille.getcellbyposition(12, iFila - 1).setFormula(Format(lnTotalPAGO, "####,###.#0"))
                Call Feuille.getcellbyposition(13, iFila - 1).setFormula(Format(lnTotalSALDO, "####,###.#0"))
            Else
                oWorkSheet.Cells(iFila, 2).Value = "Total: "
                oWorkSheet.Cells(iFila, 7).Value = Format(lnTotal, "####,###.#0")
                oWorkSheet.Cells(iFila, 8).Value = Format(lnTotalEXO, "####,###.#0")
                oWorkSheet.Cells(iFila, 9).Value = Format(lnTotalSIS, "####,###.#0")
                oWorkSheet.Cells(iFila, 10).Value = Format(lnTotalSOAT, "####,###.#0")
                oWorkSheet.Cells(iFila, 11).Value = Format(lnTotalConvenio, "####,###.#0")
                oWorkSheet.Cells(iFila, 12).Value = Format(lnTotalDEBE, "####,###.#0")
                oWorkSheet.Cells(iFila, 13).Value = Format(lnTotalPAGO, "####,###.#0")
                oWorkSheet.Cells(iFila, 14).Value = Format(lnTotalSALDO, "####,###.#0")
            End If
            If lbGeneraReciboPago = True Then
                If lnTotalPagosAdelantados > lnPagoEnFarmacia Then
                   lnTotalPagosAdelantados = lnTotalPagosAdelantados - lnPagoEnFarmacia
                   lnPagoEnFarmacia = 0
                Else
                   lnPagoEnFarmacia = lnPagoEnFarmacia - lnTotalPagosAdelantados
                   lnTotalPagosAdelantados = 0
                End If
                lnPagoEnServicios = lnPagoEnServicios - lnTotalPagosAdelantados
            Else
               Select Case mo_ReglasFacturacion.TiposFinanciamientoDevuelveComoSeTrabajaEnEstadoCuenta(lnIdTipoFinanciamientoActual, oConexion)
               Case sghTrabajaSeguroSIS
                   lnTotalSIS = lnTotalSIS - lnTotalPagosAdelantados
               Case sghTrabajaSeguroSOAT
                   lnTotalSOAT = lnTotalSOAT - lnTotalPagosAdelantados
               Case sghTrabajaSeguroConvenios
                   lnTotalConvenio = lnTotalConvenio - lnTotalPagosAdelantados
               End Select
            End If
        Next
        lnPagoEnServicios = lnTotalSALDO - lnPagoEnFarmacia
        
        iFila = iFila + 3
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(10, iFila - 1).setFormula("TOTAL CUENTA")
            Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnTotal, "####,###.#0"))
            iFila = iFila + 1
            Call Feuille.getcellbyposition(10, iFila - 1).setFormula("EXONERADO")
            Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnTotalEXO, "####,###.#0"))
            iFila = iFila + 1
            Call Feuille.getcellbyposition(10, iFila - 1).setFormula("SIS CUBRE (-PAGOS A CUENTA)")
            Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnTotalSIS, "####,###.#0"))
            iFila = iFila + 1
            Call Feuille.getcellbyposition(10, iFila - 1).setFormula("SOAT CUBRE (-PAGOS A CUENTA)")
            Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnTotalSOAT, "####,###.#0"))
            iFila = iFila + 1
            Call Feuille.getcellbyposition(10, iFila - 1).setFormula("CONVENIOS CUBRE (-PAGOS A CUENTA)")
            Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnTotalConvenio, "####,###.#0"))
            iFila = iFila + 1
            Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(10) & CStr(iFila) & ":" & mo_ReglasReportes.BuscaNombreColumna(iCol + 10) & CStr(iFila))
            mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
            Call Feuille.getcellbyposition(10, iFila - 1).setFormula("TOTAL DEUDA")
            Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnTotalDEBE, "####,###.#0"))
            iFila = iFila + 1
            Call Feuille.getcellbyposition(10, iFila - 1).setFormula("PAGOS REALIZADOS")
            Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnTotalPAGO, "####,###.#0"))
            iFila = iFila + 1
            Call Feuille.getcellbyposition(10, iFila - 1).setFormula("CAJA DEBE INGRESAR")
            iFila = iFila + 1
            Call Feuille.getcellbyposition(10, iFila - 1).setFormula("CREDITO")
            Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnTotalCredito, "####,###.#0"))
            iFila = iFila + 1
            Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(10) & CStr(iFila) & ":" & mo_ReglasReportes.BuscaNombreColumna(iCol + 10) & CStr(iFila))
            mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
        Else
            oWorkSheet.Cells(iFila, 11).Value = "TOTAL CUENTA"
            oWorkSheet.Cells(iFila, 12).Value = Format(lnTotal, "####,###.#0")
            iFila = iFila + 1
            oWorkSheet.Cells(iFila, 11).Value = "EXONERADO"
            oWorkSheet.Cells(iFila, 12).Value = Format(lnTotalEXO, "####,###.#0")
            iFila = iFila + 1
            oWorkSheet.Cells(iFila, 11).Value = "SIS CUBRE (-PAGOS A CUENTA)"
            oWorkSheet.Cells(iFila, 12).Value = Format(lnTotalSIS, "####,###.#0")
            iFila = iFila + 1
            oWorkSheet.Cells(iFila, 11).Value = "SOAT CUBRE (-PAGOS A CUENTA)"
            oWorkSheet.Cells(iFila, 12).Value = Format(lnTotalSOAT, "####,###.#0")
            iFila = iFila + 1
            oWorkSheet.Cells(iFila, 11).Value = "CONVENIOS CUBRE (-PAGOS A CUENTA)"
            oWorkSheet.Cells(iFila, 12).Value = Format(lnTotalConvenio, "####,###.#0")
            iFila = iFila + 1
            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 10, iFila, iCol + 10
            oWorkSheet.Cells(iFila, 11).Value = "TOTAL DEUDA"
            oWorkSheet.Cells(iFila, 12).Value = Format(lnTotalDEBE, "####,###.#0")
            iFila = iFila + 1
            oWorkSheet.Cells(iFila, 11).Value = "PAGOS REALIZADOS"
            oWorkSheet.Cells(iFila, 12).Value = Format(lnTotalPAGO, "####,###.#0")
            iFila = iFila + 1
            oWorkSheet.Cells(iFila, 11).Value = "CAJA DEBE INGRESAR"
            iFila = iFila + 1
            oWorkSheet.Cells(iFila, 11).Value = "CREDITO"
            oWorkSheet.Cells(iFila, 12).Value = Format(lnTotalCredito, "####,###.#0")
            iFila = iFila + 1
            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 10, iFila, iCol + 10
        End If
        
        lnPagoEnServicios = CCur(txtTotalServicios.Text) + lnTotalPagarEstancia
        lnPagoEnFarmacia = CCur(txtTotalFarmacia.Text)
        lnTotalSALDO = lnPagoEnServicios + lnPagoEnFarmacia
        
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(10, iFila - 1).setFormula("PACIENTE DEBE PAGAR")
            Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnTotalSALDO, "####,###.#0"))
            iFila = iFila + 1
            Call Feuille.getcellbyposition(10, iFila - 1).setFormula("PAGO POR CONSUMO FARMACIA")
            Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnPagoEnFarmacia, "####,###.#0"))
            iFila = iFila + 1
            Call Feuille.getcellbyposition(10, iFila - 1).setFormula("PAGO POR CONSUMO SERVICIO")
            Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnPagoEnServicios, "####,###.#0"))
        Else
            oWorkSheet.Cells(iFila, 11).Value = "PACIENTE DEBE PAGAR"
            oWorkSheet.Cells(iFila, 12).Value = Format(lnTotalSALDO, "####,###.#0")
            iFila = iFila + 1
            oWorkSheet.Cells(iFila, 11).Value = "PAGO POR CONSUMO FARMACIA"
            oWorkSheet.Cells(iFila, 12).Value = Format(lnPagoEnFarmacia, "####,###.#0")
            iFila = iFila + 1
            oWorkSheet.Cells(iFila, 11).Value = "PAGO POR CONSUMO SERVICIO"
            oWorkSheet.Cells(iFila, 12).Value = Format(lnPagoEnServicios, "####,###.#0")
        End If
        
        'Transferencias
        rsReporte.Close
        Set rsReporte = mo_ReglasFacturacion.AtencionesFiltraTransferenciasPorIdAtencion(ml_idAtencion)
        If rsReporte.RecordCount > 0 Then
            iFila = iFila - rsReporte.RecordCount - 1
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("ESTADIA")
            Else
                oWorkSheet.Cells(iFila, 2).Value = "ESTADIA"
            End If
            iFila = iFila + 1
            If lbEsOpenOffice = True Then
                Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(2) & CStr(iFila) & ":" & mo_ReglasReportes.BuscaNombreColumna(6) & CStr(iFila))
                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
            Else
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 6
            End If
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Cod.Cama")
                Call Feuille.getcellbyposition(2, iFila - 1).setFormula("Servicio que fue transferido")
                Call Feuille.getcellbyposition(4, iFila - 1).setFormula("F.Transf")
                Call Feuille.getcellbyposition(5, iFila - 1).setFormula("H.Transf")
            Else
                oWorkSheet.Cells(iFila, 2).Value = "Cod.Cama"
                oWorkSheet.Cells(iFila, 3).Value = "Servicio que fue transferido"
                oWorkSheet.Cells(iFila, 5).Value = "F.Transf"
                oWorkSheet.Cells(iFila, 6).Value = "H.Transf"
            End If
            iFila = iFila + 1
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF
                If lbEsOpenOffice = True Then
                    'Call Feuille.getcellbyposition(1, iFila - 1).setFormula(rsReporte!CodigoCama)
                    'Yamill palomino 15/10/2014
                    Call Feuille.getcellbyposition(1, iFila - 1).setFormula(IIf(IsNull(rsReporte!CodigoCama), "", (rsReporte!CodigoCama)))
                    Call Feuille.getcellbyposition(2, iFila - 1).setFormula(rsReporte!NombreServicio)
                    Call Feuille.getcellbyposition(4, iFila - 1).setFormula("'" & rsReporte!FechaOcupacion)
                    Call Feuille.getcellbyposition(5, iFila - 1).setFormula(rsReporte!HoraOcupacion)
                Else
                    oWorkSheet.Cells(iFila, 2).Value = rsReporte!CodigoCama
                    oWorkSheet.Cells(iFila, 3).Value = rsReporte!NombreServicio
                    oWorkSheet.Cells(iFila, 5).Value = "'" & rsReporte!FechaOcupacion
                    oWorkSheet.Cells(iFila, 6).Value = rsReporte!HoraOcupacion
                End If
                iFila = iFila + 1
                rsReporte.MoveNext
            Loop
        End If
        '***Donaciones en Farmacia
        If rsItemsDonaciones.RecordCount > 0 Then
           Dim lnCantidadDona As Long, lcCodigoDona As String, lbContinua As Boolean
           iFila = iFila + 2
           If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("LISTA DE DONACIONES:")
                iFila = iFila + 1
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Descripción")
                Call Feuille.getcellbyposition(8, iFila - 1).setFormula("Cantidad")
                Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(2) & CStr(iFila) & ":" & mo_ReglasReportes.BuscaNombreColumna(8) & CStr(iFila))
                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
           Else
                oWorkSheet.Cells(iFila, 2).Value = "LISTA DE DONACIONES:"
                iFila = iFila + 1
                oWorkSheet.Cells(iFila, 2).Value = "Descripción"
                oWorkSheet.Cells(iFila, 9).Value = "Cantidad"
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 9
           End If
           Set rsReporte = Nothing
           With rsReporte
                  .Fields.Append "Codigo", adVarChar, 10, adFldIsNullable
                  .Fields.Append "Nombre", adVarChar, 200, adFldIsNullable
                  .Fields.Append "Cantidad", adInteger
                  .CursorType = adOpenDynamic
                  .LockType = adLockOptimistic
                  .Open
           End With
           rsItemsDonaciones.MoveFirst
           Do While Not rsItemsDonaciones.EOF
              lbContinua = True
              If rsReporte.RecordCount > 0 Then
                 rsReporte.MoveFirst
                 rsReporte.Find "Codigo='" & rsItemsDonaciones.Fields!Codigo & "'"
                 If Not rsReporte.EOF Then
                    lbContinua = False
                 End If
              End If
              If lbContinua = True Then
                  rsReporte.AddNew
                  rsReporte.Fields!Codigo = rsItemsDonaciones.Fields!Codigo
                  rsReporte.Fields!nombre = rsItemsDonaciones.Fields!nombre
              End If
              rsReporte.Fields!cantidad = rsReporte.Fields!cantidad + rsItemsDonaciones.Fields!cantidad
              rsReporte.Update
              rsItemsDonaciones.MoveNext
           Loop
           rsReporte.Sort = "nombre,codigo"
           rsReporte.MoveFirst
           Do While Not rsReporte.EOF
              iFila = iFila + 1
              lcCodigoDona = rsReporte.Fields!Codigo
              If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula(Trim(rsReporte.Fields!Codigo) & " " & rsReporte.Fields!nombre)
                Call Feuille.getcellbyposition(8, iFila - 1).setFormula(rsReporte.Fields!cantidad)
              Else
                oWorkSheet.Cells(iFila, 2).Value = Trim(rsReporte.Fields!Codigo) & " " & rsReporte.Fields!nombre
                oWorkSheet.Cells(iFila, 9).Value = rsReporte.Fields!cantidad
              End If
              rsReporte.MoveNext
           Loop
        End If
        '
        If lcListaDeOrdenesDePago <> "" Then
            iFila = iFila + 2
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("* El CAJERO debe emitir Boletas usando " & lcListaDeOrdenesDePago)
            Else
                oWorkSheet.Cells(iFila, 2).Value = "* El CAJERO debe emitir Boletas usando " & lcListaDeOrdenesDePago
            End If
        End If
        '
        ''Yamill palomino 15/10/2014
        If lbEsOpenOffice = True Then
            'MsgBox "Se generó en forma correcta el reporte: " & lcArchivoExcel, vbInformation
            Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
            PrintArea(0).Sheet = 0
            PrintArea(0).startcolumn = 1
            PrintArea(0).StartRow = 0
            PrintArea(0).EndColumn = 17
            PrintArea(0).EndRow = iFila
'            Call Feuille.SetPrintAreas(PrintArea())
'            Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
            Call Feuille.SetPrintAreas(PrintArea())
            Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
            MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
        Else
            oWorkSheet.PageSetup.PrintTitleRows = "$1:$7"
            If oWorkSheet.PageSetup.PrintArea <> "" Then oWorkSheet.PageSetup.PrintArea = "$A$1:$R$" & (iFila + 2) 'sighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
            oExcel.Visible = True
            oWorkSheet.PrintPreview
        End If
        
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
        'encabezado de pagina
        Set PageStyles = Nothing
        Set Sheet = Nothing
        Set StyleFamilies = Nothing
        Set DefPage = Nothing
        Set Htext = Nothing
        Set Hcontent = Nothing
    Else
        'liberar memoria
        Set oExcel = Nothing
        Set oWorkBookPlantilla = Nothing
        Set oWorkBook = Nothing
        Set oWorkSheet = Nothing
    End If
    
    Set oGenerarRecordsetProductos = Nothing
    oConexion.Close
    Set oConexion = Nothing
    MousePointer = 1

End Sub

'***************daniel barrantes**************
'***************Retorna todos los ADELANTOS DE PAGO de la Cuenta Actual
'***************
Function DevuelveTotalDctosPorIdCuentaAtencion(lnIdCuentaAtencion As Long) As Double
        Dim lnDctos As Double
        Dim rsTmpDcto As New ADODB.Recordset
        Set rsTmpDcto = mo_AdminCaja.CajaComprobantesSeleccionarPorCuentaAtencion(lnIdCuentaAtencion)
        lnDctos = 0
        If rsTmpDcto.RecordCount > 0 Then
           rsTmpDcto.MoveFirst
           Do While Not rsTmpDcto.EOF
              lnDctos = lnDctos + IIf(IsNull(rsTmpDcto.Fields!Adelantos), 0, rsTmpDcto.Fields!Adelantos)
              rsTmpDcto.MoveNext
           Loop
        End If
        rsTmpDcto.Close
        Set rsTmpDcto = Nothing
        DevuelveTotalDctosPorIdCuentaAtencion = lnDctos
End Function

Private Sub btnLeerProductos_Click()
Dim oPaciente As New doPaciente
Dim rsRespuesta As New Recordset
    If txtNroOrdenPagoS.Text <> "" Then
       Exit Sub
    End If
    If txtDctoExoneracionFarm.Text <> "" Then
       Exit Sub
    End If
    If txtNroHistoria = "" Then
        Exit Sub
    End If
    If txtNroCuenta.Text <> "" Then
       Exit Sub
    End If
    MousePointer = 11
    If oConexionConsulta.State = 0 Then
      oConexionConsulta.Open sighentidades.CadenaConexion
      oConexionConsulta.CursorLocation = adUseClient
      oConexionConsulta.CommandTimeout = 150
    End If
    
    LimparDatos
    
    oPaciente.NroHistoriaClinica = Val(txtNroHistoria)
    oPaciente.IdTipoNumeracion = Val(mo_cmbIdTipoGenHistoriaClinica.BoundText)
    Set rsRespuesta = mo_AdminAdmision.PacientesFiltrar(oPaciente, False, False, "")
    rsRespuesta.Filter = "idTipoNumeracion=" & oPaciente.IdTipoNumeracion
    If rsRespuesta.RecordCount = 0 Then
        MsgBox "No se encontraron datos", vbInformation, "Búsqueda"
    ElseIf rsRespuesta.RecordCount = 1 Then
        btnImprimir.Enabled = True
        cmdImprimeCtaPorServicioHosp.Enabled = True
        bntLiquidacion.Enabled = True
        ml_idPaciente = rsRespuesta!idPaciente
'        txtPaciente.Text = rsRespuesta!ApellidoPaterno + " " + rsRespuesta!ApellidoMaterno + " " + rsRespuesta!PrimerNombre
      
        txtPaciente.Text = rsRespuesta!ApellidoPaterno + " " + rsRespuesta!ApellidoMaterno + " " + rsRespuesta!PrimerNombre + " " + _
                            IIf(IsNull(rsRespuesta!SegundoNombre), "", " " + rsRespuesta!SegundoNombre)
                            
        Dim rs As New Recordset
        ml_idEstadoCuentaAtencion = 0
        oPaciente.idPaciente = ml_idPaciente
        Set rs = mo_AdminAdmision.AtencionesBusquedaDeAtenciones(oPaciente)
        If rs.RecordCount = 1 Then
            ml_idCuentaAtencion = rs!idCuentaAtencion
            txtEstadoCuenta.Text = rs!EstadoCuenta
            ml_idAtencion = rs.Fields!idAtencion
            ml_IdTipoServicio = rs.Fields!idTipoServicio
            ml_idEstadoCuentaAtencion = rs.Fields!IdEstado
        ElseIf rs.RecordCount > 1 Then
            Dim busqueda As New BusquedaAtenciones
            Set busqueda.Atenciones = rs
            busqueda.Show 1
            txtEstadoCuenta.Text = busqueda.EstadoCuenta
            ml_idCuentaAtencion = busqueda.idCuentaAtencion
            ml_idAtencion = rs.Fields!idAtencion
            ml_IdTipoServicio = rs.Fields!idTipoServicio
            ml_idEstadoCuentaAtencion = rs.Fields!IdEstado
            Unload busqueda
        Else
            MsgBox "El paciente no tiene cuentas de atenciones en estado ABIERTO", vbInformation, "Estado de cuenta"
            btnImprimir.Enabled = False
            cmdImprimeCtaPorServicioHosp.Enabled = False
            rs.Close
            MousePointer = 1
            Exit Sub
        End If
        rs.Close
        EncontroCuenta
    End If
    MousePointer = 1
End Sub

Sub EncontroCuenta()
          On Error GoTo errEncCta       'debb-02/05/2016
          Dim oRs As New Recordset
          'carga datos de PLAN ATENCION
10        lnIdPlanActual = 0
20        If ml_idAtencion > 0 Then
30           Set oRs = mo_AdminAdmision.AtencionesXidAtencionCondicionAlta(ml_idAtencion, ml_dCondicionAlta, oConexionConsulta)
40           If oRs.RecordCount > 0 Then
50              txtEstadoCuenta.Text = "(" & Trim(txtEstadoCuenta.Text) & ") IAFA: " & Left(oRs.Fields!Descripcion, 15) & "  (PP=" & lcdTipoFinanciamiento & ")"
60              lnIdPlanActual = IIf(IsNull(oRs.Fields!IdFuenteFinanciamiento), 0, oRs.Fields!IdFuenteFinanciamiento)
70              If ml_dCondicionAlta <> "" Then
80                 ml_dCondicionAlta = "(Cond.AM: " & Trim(oRs.Fields!dCondicionAlta) & ")"
90              End If
100          End If
110          oRs.Close
120       End If
130       Set oRs = Nothing
          'Deudas Anteriores
140       ucMensajeParpadeando1.MensajeDeTexto = mo_ReglasFacturacion.DevuelveDeudaPacienteDeAntencionesAnteriores(ml_idPaciente, oConexionConsulta, ml_idCuentaAtencion)
150       If ucMensajeParpadeando1.MensajeDeTexto <> "" Then
160          ucMensajeParpadeando1.MensajeDeTexto = "Deudas: " & ucMensajeParpadeando1.MensajeDeTexto
170       End If
180       ucMensajeParpadeando1.Visible = True
          '
190       CargaCuentaElegida
          '
200       CargaCtasDelPaciente
          '
210       CargaGridReembolsos
          '
220       CargaGrillaDonaciones
    
Exit Sub                'debb-02/05/2016
errEncCta:              'debb-02/05/2016
          MsgBox Err.Number & " " & Err.Description & _
          sighentidades.DevuelveFuenteDeLineaDelError(Erl(), "Sub EncontroCuenta", "ucEstadoCuenta.ctl")   'debb-02/05/2016

End Sub



Sub CargaCuentaElegida()
              On Error GoTo ErrCargaCta                 'debb-02/05/2016
10            ObtenerDatosDeLaCuentaDeAtencion
20            cmdExoneracion.Visible = False
30            btnAceptar.Visible = False
40            btnAceptar.Enabled = True
50            bntLiquidacion.Visible = False
60            chkServiciosTodos.Visible = False
70            chkFarmaciaTodos.Visible = False
              '
80            lbTieneAccesoActualizar = True
90            If Not (((ml_idEstadoCuentaAtencion = sghConAltaMedica Or ml_idEstadoCuentaAtencion = sghAbierto) And ml_IdTipoServicio <> sghConsultaExterna) Or (ml_idEstadoCuentaAtencion = sghAbierto And ml_IdTipoServicio = sghConsultaExterna)) Then
100              MsgBox "Fijese el ESTADO de esta N° Cuenta", vbInformation, "Mensaje"
110              lbTieneAccesoActualizar = False
120           End If
              '
130           btnCtaPagada.Visible = False
140           btnPendientePagoSeguro.Visible = False
150           btnAbrirCuenta.Visible = False
160           btnCerrarCuenta.Visible = False
170           btnCtaAnulada.Visible = False
180           btnCtaGarante.Visible = False
190           Select Case ml_idUsuarioConPermisoEnSISoEXOoSOAT
              Case 2     'SIS
200               If mo_ReglasFacturacion.TiposFinanciamientoDevuelveComoSeTrabajaEnEstadoCuenta(lnIdTipoFinanciamientoActual, oConexionConsulta) = ml_idUsuarioConPermisoEnSISoEXOoSOAT Then
210                  btnAbrirCuenta.Visible = True
220               Else
230                  lbTieneAccesoActualizar = False
240               End If
250           Case 3     'SOAT
260               If mo_ReglasFacturacion.TiposFinanciamientoDevuelveComoSeTrabajaEnEstadoCuenta(lnIdTipoFinanciamientoActual, oConexionConsulta) = ml_idUsuarioConPermisoEnSISoEXOoSOAT Then
270                  btnAbrirCuenta.Visible = True
280               Else
290                  lbTieneAccesoActualizar = False
300               End If
310           Case 4     'Convenio FOSPOLIS
320               If mo_ReglasFacturacion.TiposFinanciamientoDevuelveComoSeTrabajaEnEstadoCuenta(lnIdTipoFinanciamientoActual, oConexionConsulta) = ml_idUsuarioConPermisoEnSISoEXOoSOAT Then
330                  btnAbrirCuenta.Visible = True
340               Else
350                  lbTieneAccesoActualizar = False
360               End If
370           Case 9     'Exoneraciones
380               If mo_ReglasFacturacion.TiposFinanciamientoDevuelveComoSeTrabajaEnEstadoCuenta(lnIdTipoFinanciamientoActual, oConexionConsulta) = ml_idUsuarioConPermisoEnSISoEXOoSOAT Or mo_ReglasFacturacion.TiposFinanciamientoDevuelveComoSeTrabajaEnEstadoCuenta(lnIdTipoFinanciamientoActual, oConexionConsulta) = 1 Then
390                  btnAbrirCuenta.Visible = True
400               Else
410                  lbTieneAccesoActualizar = False
420               End If
430               cmdExoneracion.Visible = True
440               cmdExoneracion.Enabled = True
450           End Select
460           If mo_ReglasSeguridad.TieneRolAdministrador(ml_idUsuario) Then
470               btnCtaPagada.Visible = True
480               btnPendientePagoSeguro.Visible = True
490               btnAbrirCuenta.Visible = True
500               btnCerrarCuenta.Visible = True
510               btnCtaAnulada.Visible = True
520               btnCtaGarante.Visible = True
                  'bntLiquidacion.Visible = True
530           End If
540           If lbTieneAccesoActualizar = True Then
550               btnCtaPagada.Visible = True
560               btnPendientePagoSeguro.Visible = True
570               btnAbrirCuenta.Visible = True
580               btnCerrarCuenta.Visible = True
590               btnCtaAnulada.Visible = True
600               btnCtaGarante.Visible = True
610               Select Case ml_idUsuarioConPermisoEnSISoEXOoSOAT
                  Case 1000     'Solo tiene opcion a CONSULTAR
620                  btnAceptar.Visible = False
630               Case 2     'SIS
640                   btnAceptar.Visible = True
650                   btnAceptar.Caption = "Actualiza SIS"
660                   bntLiquidacion.Visible = True
670                   chkServiciosTodos.Visible = True
680                   chkFarmaciaTodos.Visible = True
690                   If ml_IdTipoServicio = sghConsultaExterna Then     'Solo para CONSULTA EXTERNA
                         'cmdNewServicios.Visible = True
                         'cmdNewFarmacia.Visible = True
700                   End If
710               Case 3     'SOAT
720                   btnAceptar.Visible = True
730                   btnAceptar.Caption = "Actualiza SOAT"
740                   bntLiquidacion.Visible = True
750                   chkServiciosTodos.Visible = True
760                   chkFarmaciaTodos.Visible = True
770                   If ml_IdTipoServicio = sghConsultaExterna Then    'Solo para CONSULTA EXTERNA
780                   End If
790               Case 4     'Convenio FOSPOLIS
800                   btnAceptar.Visible = True
810                   btnAceptar.Caption = "Actualiza CONVENIO FOSPOLIS"
820                   bntLiquidacion.Visible = True
830                   If ml_IdTipoServicio = sghConsultaExterna Then     'Solo para CONSULTA EXTERNA
840                   End If
850               Case 9     'Exoneraciones
860                   btnAceptar.Visible = True
870                   btnAceptar.Caption = "Actualiza Exoneraciones"
880                   bntLiquidacion.Visible = True
890                   chkFarmaciaTodos.Visible = True
900                   chkServiciosTodos.Visible = True
910               End Select
920           End If
              'debb-01/03/2016 (inicio)
930           If ml_lbPuedeVerResultados = True Then
940              btnCtaPagada.Visible = True
950              btnPendientePagoSeguro.Visible = True
960              btnAbrirCuenta.Visible = True
970              btnCerrarCuenta.Visible = True
980              btnCtaAnulada.Visible = True
990              btnCtaGarante.Visible = True
1000             btnRecalculaPlan.Visible = True
1010             btnCtaPagada.Enabled = True
1020             btnPendientePagoSeguro.Enabled = True
1030             btnAbrirCuenta.Enabled = True
1040             btnCerrarCuenta.Enabled = True
1050             btnCtaAnulada.Enabled = True
1060             btnCtaGarante.Enabled = True
1070             btnRecalculaPlan.Enabled = True
1080          End If
              'debb-01/03/2016 (fin)
              Exit Sub              'debb-02/05/2016
ErrCargaCta:                        'debb-02/05/2016
          MsgBox Err.Number & " " & Err.Description & _
          sighentidades.DevuelveFuenteDeLineaDelError(Erl(), "Sub CargaCuentaElegida", "ucEstadoCuenta.ctl")   'debb-02/05/2016
              
End Sub

Sub LimparDatos()
    ucFacturacionServicios.LimpiarGrilla
    ucFacturacionBienes.LimpiarGrilla
    Set grdCabecera.DataSource = Nothing
    txtFingreso.Text = ""
    txtFegreso.Text = ""
    txtCuenta.Text = ""
    txtEstadoCuenta.Text = ""
    txtPaciente.Text = ""
    txtDevoluciones.Text = ""
    txtTotalServicios.Text = ""
    txtTotalFarmacia.Text = ""
    txtServicio.Text = ""
    chkServiciosTodos.Value = 0
    chkFarmaciaTodos.Value = 0
    cmbFuenteFinanciamiento.Text = ""
    txtTotalSeguroFarmacia.Text = "0"
    txtTotalSeguroServicio.Text = "0"
    txtNroCuenta.Text = ""
    txtNroOrdenPagoS.Text = ""
    ml_idAtencion = 0
    txtDomicilioPacienteEnAtencion.Text = ""
    txtFAltaAdm.Text = ""
    txtFapertura.Text = ""
    ucMensajeParpadeando1.MensajeDeTexto = ""
    txtDxEgr.Text = ""
    txtDctoExoneracionFarm.Text = ""
    lnTotalPagosAdelantados = 0
    txtPagosAdelantoS.Text = ""
    txtPagosAdelantoF.Text = ""
    txtPagosAdelantoC.Text = ""
    btnImprimir.Enabled = False
    cmdImprimeCtaPorServicioHosp.Enabled = False
    bntLiquidacion.Enabled = False
    txtReembolsoF.Text = ""
    txtReembolsoS.Text = ""
    txtReembolsoT.Text = ""
    txtPorReembolsar.Text = ""
    lcdTipoFinanciamiento = ""
    ml_idEstadoCuentaAtencion = 0
    cmbFormaPago.Text = ""
    lcListaDeOrdenesDePago = ""
    ml_lbEsPacienteExterno = False
    lnPagosXdevoluciones = 0
    Set rsItemsDonaciones = Nothing
    Set grdItemsDonaciones.DataSource = Nothing
End Sub
Sub ObtenerDatosDeLaCuentaDeAtencion()
10            On Error GoTo errObt                  'debb-02/05/2016
20            If ml_idCuentaAtencion <> 0 Then
                  Dim lnTotalPagoSeguro As Double, lnTotalPagoDelPaciente As Double, lnTotalizaPagosDelPacienteConSeguro As Double
                  Dim lnIdTipoConceptoFarmacia As Integer, lnTotalApagarS As Double, lnTotalApagarF As Double
30                Set mo_DOCuentaAtencion = mo_ReglasFacturacion.CuentasAtencionSeleccionarPorId(ml_idCuentaAtencion, oConexionConsulta)
40                txtFapertura.Text = Format(mo_DOCuentaAtencion.fechacreacion, sighentidades.DevuelveFechaSoloFormato_DMY_HM)
                  
                  'LEER DATOS DE LA ATENCION
50                LeerDatosAtencion
60                lbGeneraReciboPago = mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(lnIdTipoFinanciamientoActual, oConexionConsulta)
70                lnEstadoFacturacionAtendidoOpreventa = sghAtendido
80                lnIdTipoConceptoFarmacia = mo_ReglasFacturacion.FuentesFinanciamientoDevuelveIdTipoConceptoFarmacia(lnIdPlanActual, oConexionConsulta)
90                CreaTemporales
                  
                  'LEER DATOS DE SERVICIOS
100               ucFacturacionServicios.ProcesoEnElServidor = mb_ProcesoEnElServidor
                  
110               ucFacturacionServicios.LimpiarGrilla
120               ucFacturacionServicios.PuedeVerResultados = ml_lbPuedeVerResultados    'Ver Resultado de Laboratorio
130               ucFacturacionServicios.idPaciente = ml_idPaciente                      'Ver Resultado de Laboratorio
140               ucFacturacionServicios.Paciente = txtPaciente.Text                     'Ver Resultado de Laboratorio
150               ucFacturacionServicios.idTipoSexo = ml_idTipoSexo                      'Ver Resultado de Laboratorio
160               ucFacturacionServicios.EstadosFacturacion = ""
170               ucFacturacionServicios.idTipoFinanciamiento = lnIdTipoFinanciamientoActual
180               ucFacturacionServicios.TipoProducto = sghServicio
190               ucFacturacionServicios.idCuentaAtencion = Val(ml_idCuentaAtencion)
200               ucFacturacionServicios.AgruparPor = Val(cmbAgrupar.ItemData(cmbAgrupar.ListIndex))
210               ucFacturacionServicios.CargaProductosPorIdCuentaAtencion lnTotalPagoSeguro, lnTotalPagoDelPaciente, lnTotalizaPagosDelPacienteConSeguro, oRsCuentaCabecera, oRsCuentaDetalle, lnIdTipoConceptoFarmacia, lnTotalApagarS
220               ucFacturacionServicios.ActualizaPreciosImportesEnTodosItemsParaSisSoat (ml_idUsuarioConPermisoEnSISoEXOoSOAT)
230               txtTotalServicios.Text = lnTotalPagoDelPaciente
240               txtTotalSeguroServicio.Text = lnTotalPagoSeguro
250               If lbGeneraReciboPago = True Then
260                     txtTotalServicios.Text = lnTotalizaPagosDelPacienteConSeguro
270               End If
                  
                  'LEER DATOS DE BIENES E INSUMOS
280               ucFacturacionBienes.ProcesoEnElServidor = mb_ProcesoEnElServidor
290               ucFacturacionBienes.LimpiarGrilla
300               ucFacturacionBienes.EstadosFacturacion = ""
310               ucFacturacionBienes.idTipoFinanciamiento = lnIdTipoFinanciamientoActual
320               ucFacturacionBienes.TipoProducto = sghbien
330               ucFacturacionBienes.idCuentaAtencion = Val(ml_idCuentaAtencion)
340               ucFacturacionBienes.AgruparPor = Val(cmbAgrupar.ItemData(cmbAgrupar.ListIndex))
350               ucFacturacionBienes.CargaProductosPorIdCuentaAtencion lnTotalPagoSeguro, lnTotalPagoDelPaciente, lnTotalizaPagosDelPacienteConSeguro, oRsCuentaCabecera, oRsCuentaDetalle, lnIdTipoConceptoFarmacia, lnTotalApagarF
360               ucFacturacionBienes.ActualizaPreciosImportesEnTodosItemsParaSisSoat (ml_idUsuarioConPermisoEnSISoEXOoSOAT)
370               txtTotalFarmacia.Text = lnTotalPagoDelPaciente
380               txtTotalSeguroFarmacia.Text = lnTotalPagoSeguro
390               If lbGeneraReciboPago = True Then
400                   txtTotalFarmacia.Text = lnTotalizaPagosDelPacienteConSeguro
410               End If
                  
                  '***************** GalenHos v.3.0 (inicio)*****************
                  '*********** Pagos Adelantados  (inicio)
                  '*********** Se da preferencia a CANCELAR primero la deuda de FARMACIA
                  '*********** el Resto que queda será para CANCELAR parte o todo SERVICIOS
420               lnTotalPagosAdelantados = mo_AdminCaja.RetornaTotalDescuentosPorAdelantosSegunCuenta(ml_idCuentaAtencion, oConexionConsulta)
430               lnPagosXdevoluciones = mo_ReglasFacturacion.RetornaImporteDePagosXdevolucionesPorNroCuenta(ml_idCuentaAtencion, oConexionConsulta)
440               txtDevoluciones.Text = lnPagosXdevoluciones
450               txtPagosAdelantoC.Text = Val(Trim(Str(lnTotalPagosAdelantados))) - Val(Trim(Str((lnPagosXdevoluciones))))
460               lnTotalPagosAdelantados = Val(txtPagosAdelantoC.Text)
470               If lbGeneraReciboPago = True Then
480                   If ml_lbEsPacienteExterno = True Then
490                     lcListaDeOrdenesDePago = mo_ReglasFacturacion.DevuelveOrdenesPagoSegunCuenta(ml_idCuentaAtencion, oConexionConsulta)
500                   End If
510                   txtTotalSeguroFarmacia.Visible = False
520                   lblTotalSeguroFarmacia.Visible = False
530                   txtTotalSeguroServicio.Visible = False
540                   lblTotalSeguroServicio.Visible = False
550                   txtTotalFarmacia.Visible = True
560                   lblPagoFarmacia.Visible = True
570                   txtTotalServicios.Visible = True
580                   lblTotalServicios.Visible = True
                      'Contado
590                   If lnTotalPagosAdelantados > Val(txtTotalFarmacia.Text) Then
600                      lnTotalPagosAdelantados = lnTotalPagosAdelantados - Val(txtTotalFarmacia.Text)
610                      txtPagosAdelantoF.Text = txtTotalFarmacia.Text
620                      txtTotalFarmacia.Text = 0
630                   Else
640                      txtPagosAdelantoF.Text = lnTotalPagosAdelantados
650                      txtTotalFarmacia.Text = Val(txtTotalFarmacia.Text) - lnTotalPagosAdelantados
660                      lnTotalPagosAdelantados = 0
670                   End If
680                   If lnTotalPagosAdelantados > 0 Then
690                          txtTotalServicios.Text = txtTotalServicios.Text - lnTotalPagosAdelantados
700                          txtPagosAdelantoS.Text = lnTotalPagosAdelantados
710                   End If
720               Else
                     'Seguros
730                   txtTotalSeguroFarmacia.Visible = True
740                   lblTotalSeguroFarmacia.Visible = True
750                   txtTotalSeguroServicio.Visible = True
760                   lblTotalSeguroServicio.Visible = True
770                   txtTotalFarmacia.Visible = True
780                   lblPagoFarmacia.Visible = True
790                   txtTotalServicios.Visible = True
800                   lblTotalServicios.Visible = True
810                   If lnTotalPagosAdelantados > Val(txtTotalSeguroFarmacia.Text) Then
820                      lnTotalPagosAdelantados = lnTotalPagosAdelantados - Val(txtTotalSeguroFarmacia.Text)
830                      txtPagosAdelantoF.Text = txtTotalSeguroFarmacia.Text
840                      txtTotalSeguroFarmacia.Text = 0
850                   Else
860                      txtPagosAdelantoF.Text = lnTotalPagosAdelantados
870                      txtTotalSeguroFarmacia.Text = Val(txtTotalSeguroFarmacia.Text) - lnTotalPagosAdelantados
880                      lnTotalPagosAdelantados = 0
890                   End If
900                   If lnTotalPagosAdelantados > 0 Then
910                       If lnTotalPagosAdelantados > Val(txtTotalSeguroServicio.Text) Then
920                          txtPagosAdelantoS.Text = txtTotalSeguroServicio.Text
930                          txtTotalSeguroServicio.Text = lnTotalPagosAdelantados - Val(txtTotalSeguroServicio.Text)
940                       Else
950                          txtPagosAdelantoS.Text = lnTotalPagosAdelantados
960                          txtTotalSeguroServicio.Text = txtTotalSeguroServicio.Text - lnTotalPagosAdelantados
970                       End If
980                   End If
990               End If
                  '*********** Pagos Adelantados  (fin)
                  '***************** GalenHos v.3.0 (fin)*****************
1000              If mo_PermisosFacturacion.AbrirCuentaAtencion Then
1010                  btnAbrirCuenta.Enabled = True
1020              End If
1030              If mo_PermisosFacturacion.CerrarCuentaAtencion Then
1040                  btnCerrarCuenta.Enabled = True
1050              End If
1060              txtTotalApagar.Text = lnTotalApagarS + lnTotalApagarF - Val(txtPagosAdelantoC.Text)
1070              oRsCuentaCabecera.Sort = "fecha"
1080              Set grdCabecera.DataSource = oRsCuentaCabecera
1090              Set grdDetalle.DataSource = Nothing
1100          Else
1110              MsgBox "Por favor seleccione una cuenta de atención", vbInformation, "Búsqueda de cuentas de atención"
1120          End If
1130          Exit Sub    'debb-02/05/2016
errObt:             'debb-02/05/2016
1140  MsgBox Err.Description & _
      sighentidades.DevuelveFuenteDeLineaDelError(Erl(), "Sub ObtenerDatosDeLaCuentaDeAtencion", "ucEstadoCuenta.ctl")          'debb-02/05/2016
End Sub

Sub LeerDatosAtencion()
10        On Error GoTo errLeerDA
          Dim oRsTmp As New ADODB.Recordset
          Dim oDODiagnostico As New DODiagnostico
20        If ml_IdTipoServicio = sghConsultaExterna Or ml_lbEsPacienteExterno = True Then
30           Set oRsTmp = mo_ReglasFacturacion.AtencionesFiltraIngresosPorIdAtencion(ml_idAtencion)
40        Else
50           Set oRsTmp = mo_ReglasFacturacion.AtencionesFiltraEgresosPorIdAtencion(ml_idAtencion)
60        End If
70        If oRsTmp.RecordCount > 0 Then
80           txtFingreso.Text = Format(oRsTmp.Fields!FechaIngreso & " " & oRsTmp.Fields!HoraIngreso, sighentidades.DevuelveFechaSoloFormato_DMY_HM)
90           txtFegreso.Text = IIf(IsNull(oRsTmp.Fields!FechaEgreso), "", Format(oRsTmp.Fields!FechaEgreso & " " & oRsTmp.Fields!HoraEgreso, sighentidades.DevuelveFechaSoloFormato_DMY_HM))
100          txtServicio.Text = IIf(IsNull(oRsTmp.Fields!DServicio), "", Trim(oRsTmp.Fields!DServicio)) & " (" & mo_ReglasFacturacion.DevuelveDescripcionTipoServicioSegunIdTipoServicio(oRsTmp.Fields!idTipoServicio) & ") " & IIf(IsNull(oRsTmp.Fields!codCama), "", "  Cod.Cama: " & oRsTmp.Fields!codCama) & " " & ml_dCondicionAlta
110          txtDomicilioPacienteEnAtencion.Text = IIf(IsNull(oRsTmp.Fields!DireccionDomicilio), "", oRsTmp.Fields!DireccionDomicilio)
120          lnIdTipoFinanciamientoActual = oRsTmp.Fields!IdFormaPago
130          txtFAltaAdm.Text = IIf(IsNull(oRsTmp.Fields!FechaEgresoAdministrativo), "", Format(oRsTmp.Fields!FechaEgresoAdministrativo & " " & oRsTmp.Fields!HoraEgresoAdministrativo, sighentidades.DevuelveFechaSoloFormato_DMY_HM))
140       End If
150       txtCuenta.Text = ml_idCuentaAtencion
          '
160       Set oDODiagnostico = mo_ReglasFacturacion.DevuelveDxAltaMedicaConexion(ml_idAtencion, ml_IdTipoServicio, oConexionConsulta)
170       If oDODiagnostico.Descripcion = "" Then
180          txtDxEgr.Text = "Dx INGRESO: " & mo_ReglasFacturacion.DevuelveDxIngresoSoloHospEmerg(ml_IdTipoServicio, ml_idAtencion, oConexionConsulta)
190       Else
200          txtDxEgr.Text = Trim(oDODiagnostico.CodigoCIE2004) & " " & oDODiagnostico.Descripcion
210       End If
220       Set oDODiagnostico = Nothing
230       Exit Sub    'debb-02/05/2016
errLeerDA:      'debb-02/05/2016
240      MsgBox Err.Description & _
                   sighentidades.DevuelveFuenteDeLineaDelError(Erl(), "Sub LeerDatosAtencion", "ucEstadoCuenta.ctl")          'debb-02/05/2016

End Sub


Public Sub inicializar()
    mb_ProcesoEnElServidor = True
    InicilizarParametros

    Set mo_cmbIdTipoGenHistoriaClinica.MiComboBox = cmbIdTipoGenHistoriaClinica
    Set mo_cmbFechaIngreso.MiComboBox = cmbFechaIngreso
    
    ConfigurarFechaIngreso
    ConfigurarTiposHistoriaClinica
    ConfigurarComboAgrupar
    UsuarioConPermisoEnSISoEXOoSOAT
    CargaTextos
    
    mo_Formulario.HabilitarDeshabilitar txtPaciente, False
    mo_Formulario.HabilitarDeshabilitar cmbFechaIngreso, False
    mo_Formulario.HabilitarDeshabilitar txtFingreso, False
    mo_Formulario.HabilitarDeshabilitar txtFegreso, False
    mo_Formulario.HabilitarDeshabilitar txtEstadoCuenta, False
    mo_Formulario.HabilitarDeshabilitar txtCuenta, False
    mo_Formulario.HabilitarDeshabilitar txtServicio, False
    mo_Formulario.HabilitarDeshabilitar txtDomicilioPacienteEnAtencion, False
    mo_Formulario.HabilitarDeshabilitar txtCtaPagada, False
    mo_Formulario.HabilitarDeshabilitar txtCtaAbrir, False
    mo_Formulario.HabilitarDeshabilitar txtCtaCerrar, False
    mo_Formulario.HabilitarDeshabilitar txtCtaAnulada, False
    mo_Formulario.HabilitarDeshabilitar txtRecalculo, False
    mo_Formulario.HabilitarDeshabilitar txtPendienteSeguro, False
    mo_Formulario.HabilitarDeshabilitar txtCtaConGarante, False
    mo_Formulario.HabilitarDeshabilitar txtNroHistoria, False
    mo_Formulario.HabilitarDeshabilitar txtFAltaAdm, False
    mo_Formulario.HabilitarDeshabilitar txtFapertura, False
    mo_Formulario.HabilitarDeshabilitar txtDxEgr, False
    mo_Formulario.HabilitarDeshabilitar txtTotalApagar, False
    mo_Formulario.HabilitarDeshabilitar txtPagosAdelantoC, False
    mo_Formulario.HabilitarDeshabilitar txtTotalFarmacia, False
    mo_Formulario.HabilitarDeshabilitar txtTotalSeguroFarmacia, False
    mo_Formulario.HabilitarDeshabilitar txtPagosAdelantoF, False
    mo_Formulario.HabilitarDeshabilitar txtTotalServicios, False
    mo_Formulario.HabilitarDeshabilitar txtTotalSeguroServicio, False
    mo_Formulario.HabilitarDeshabilitar txtPagosAdelantoS, False
    mo_Formulario.HabilitarDeshabilitar txtTotalConsumo, False
    mo_Formulario.HabilitarDeshabilitar txtExoneraciones, False
    mo_Formulario.HabilitarDeshabilitar txtTotalSeguro, False
    mo_Formulario.HabilitarDeshabilitar txtDevoluciones, False
    mo_Formulario.HabilitarDeshabilitar txtTotalDonaciones, False
    mo_Formulario.HabilitarDeshabilitar txtReembolsoF, False
    mo_Formulario.HabilitarDeshabilitar txtReembolsoS, False
    mo_Formulario.HabilitarDeshabilitar txtReembolsoT, False
    mo_Formulario.HabilitarDeshabilitar txtPorReembolsar, False
    
    ucFacturacionServicios.inicializar
    ucFacturacionServicios.idUsuarioConPermisoEnSISoEXOoSOAT = ml_idUsuarioConPermisoEnSISoEXOoSOAT
    ucFacturacionBienes.inicializar
    ucFacturacionBienes.idUsuarioConPermisoEnSISoEXOoSOAT = ml_idUsuarioConPermisoEnSISoEXOoSOAT
    
    cmbAgrupar.ListIndex = 0
    mo_cmbIdTipoGenHistoriaClinica.BoundText = lcBuscaParametro.SeleccionaFilaParametro(212)
    'PERMISOS
    Dim oRsPermisos As New Recordset
    Dim lbHabilitaTabResumen As Boolean
    Set oRsPermisos = mo_ReglasSeguridad.UsuariosRolesSeleccionarPermisosTodos(ml_idUsuario)
    lbHabilitaTabResumen = False
    UserControl.btnAbrirCuenta.Enabled = False
    UserControl.btnCerrarCuenta.Enabled = False
    UserControl.btnCtaAnulada.Enabled = False
    UserControl.btnCtaPagada.Enabled = False
    UserControl.btnRecalculaPlan.Enabled = False
    UserControl.btnPendientePagoSeguro.Enabled = False
    lc_TipoFinanciamientoPermitidos = ""
    ml_lbPuedeVerResultados = False
    If oRsPermisos.RecordCount > 0 Then
       Do While Not oRsPermisos.EOF
          Select Case oRsPermisos.Fields!IdPermiso
          Case 110    'Autorizado a 'Abrir Cuenta Atencion'
               UserControl.btnAbrirCuenta.Enabled = True
               lbHabilitaTabResumen = True
          Case 111    'Autorizado a 'Cerrar Cuenta Atencion'
               UserControl.btnCerrarCuenta.Enabled = True
               lbHabilitaTabResumen = True
          Case 112    'Autorizado a 'Cuenta Anulada'
               UserControl.btnCtaAnulada.Enabled = True
               lbHabilitaTabResumen = True
          Case 113    'Autorizado a 'Cuenta Pagada'
               UserControl.btnCtaPagada.Enabled = True
               lbHabilitaTabResumen = True
          Case 114    'Autorizado a 'RECALCULOS de Cuenta'
               UserControl.btnRecalculaPlan.Enabled = True
               '********daniel-20/12/2009-inicio***********
               UserControl.btnRecalculaPlan.Visible = True
               '********daniel-20/12/2009-fin***********
               lbHabilitaTabResumen = True
          Case 115    'Autorizado a 'Cuenta Pendiente Pago Seguros'
               UserControl.btnPendientePagoSeguro.Enabled = True
               lbHabilitaTabResumen = True
          Case 116    'Autorizado a 'Cuenta con Garante'
               btnCtaGarante.Enabled = True
               lbHabilitaTabResumen = True
          Case 123    'Facturacion - Puede ver RESULTADOS en Estado Cuenta
               ml_lbPuedeVerResultados = True
          End Select
          oRsPermisos.MoveNext
       Loop
    End If
    Set oRsPermisos = Nothing
    If lbHabilitaTabResumen = False Then
       UserControl.ucFacturacionBienesInsumos.TabVisible(3) = False
    Else
       UserControl.ucFacturacionBienesInsumos.TabVisible(3) = True
    End If
    '
    gridInfra.ConfigurarFilasBiColores grdCabecera, sighentidades.GrillaConFilasBicolor
    gridInfra.ConfigurarFilasBiColores grdDetalle, sighentidades.GrillaConFilasBicolor
    gridInfra.ConfigurarFilasBiColores grdCuentasPorTipoServicio, sighentidades.GrillaConFilasBicolor
    gridInfra.ConfigurarFilasBiColores grdReembolsoF, sighentidades.GrillaConFilasBicolor
    gridInfra.ConfigurarFilasBiColores grdItemsDonaciones, sighentidades.GrillaConFilasBicolor
    '
    grdCuentasPorTipoServicio.Left = 50
    grdCuentasPorTipoServicio.Height = 7900
    '
    lnIdPagosACuenta = Val(lcBuscaParametro.SeleccionaFilaParametro(245))
    lnIdPagosXdevoluciones = Val(lcBuscaParametro.SeleccionaFilaParametro(265))
    '
    txtFechaInicio.Text = sighentidades.PrimerFechaDDMMYYDelMesActual()
    txtFechaFin.Text = Format(Date, sighentidades.DevuelveFechaSoloFormato_DMY)
End Sub

'***************daniel barrantes**************
'***************Retorna Permiso del Usuario Galenhos para poder usar solo SIS o SOAT o EXONERAR o CONVENIO_FOSPOLIS
'***************
Sub UsuarioConPermisoEnSISoEXOoSOAT()
    Dim oBuscaDondeLabora As New SIGHNegocios.ReglasComunes
    Dim oRsBuscaLabora As Recordset
    Set oRsBuscaLabora = oBuscaDondeLabora.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghSeguros, ml_idUsuario)
    If oRsBuscaLabora.RecordCount > 0 Then
       ml_idUsuarioConPermisoEnSISoEXOoSOAT = oRsBuscaLabora.Fields!idLaboraSubArea
    End If
    Set oRsBuscaLabora = Nothing
    Set oBuscaDondeLabora = Nothing
End Sub

Sub ConfigurarComboAgrupar()

        cmbAgrupar.AddItem "<Sin agrupar>":  cmbAgrupar.ItemData(cmbAgrupar.NewIndex) = "1"
        cmbAgrupar.AddItem "Nro de orden":  cmbAgrupar.ItemData(cmbAgrupar.NewIndex) = "2"
        cmbAgrupar.AddItem "Tipo de financiamiento":  cmbAgrupar.ItemData(cmbAgrupar.NewIndex) = "3"
        cmbAgrupar.AddItem "Nro atención":  cmbAgrupar.ItemData(cmbAgrupar.NewIndex) = "4"
        cmbAgrupar.AddItem "Punto de carga":  cmbAgrupar.ItemData(cmbAgrupar.NewIndex) = "5"
        '
        On Error Resume Next
        Set oRsFuentesFinanciamiento = mo_ReglasFacturacion.FuentesFinanciamientoSoloParticular
        Set cmbFuenteFinanciamiento.RowSource = oRsFuentesFinanciamiento
        cmbFuenteFinanciamiento.ListField = "Descripcion"
        cmbFuenteFinanciamiento.BoundColumn = "idFuenteFinanciamiento"
End Sub

Sub ConfigurarFechaIngreso()
    
    mo_cmbFechaIngreso.ListField = "DescripcionLarga"
    mo_cmbFechaIngreso.BoundColumn = "IdCuentaAtencion"

End Sub

Sub ConfigurarTiposHistoriaClinica()
        
        mo_cmbIdTipoGenHistoriaClinica.BoundColumn = "IdTipoNumeracion"
        mo_cmbIdTipoGenHistoriaClinica.ListField = "DescripcionLarga"
        Set mo_cmbIdTipoGenHistoriaClinica.RowSource = mo_AdminArchivoClinico.TiposGeneracionHistoriasSeleccionarTodos()

End Sub

Private Sub btnPendientePagoSeguro_Click()
    
    If Not IsDate(txtFegreso.Text) And ml_IdTipoServicio <> sghConsultaExterna Then
       MsgBox "Tiene que dar ALTA MEDICA", vbInformation, ""
       Exit Sub
    End If
    '
    If MsgBox("Esta seguro que la Cuenta pase a estado=PENDIENTE PAGO SEGURO", vbQuestion + vbYesNo, "Facturación") = vbYes Then
        Dim Login As New Login
        Login.UsuarioDeEstadoDeCuenta = ml_idUsuario
        Login.CargaDesdeOtraOpcion = True
        Login.Show vbModal
        If Not Login.Autenticado Or Login.IdUsuarioAutenticado <> ml_idUsuario Then
            Exit Sub
        End If
        mo_DOCuentaAtencion.IdUsuarioAuditoria = ml_idUsuario
        mo_DOCuentaAtencion.TotalPorPagar = CCur(txtTotalSeguroServicio.Text) + CCur(txtTotalSeguroFarmacia.Text)
        If mo_ReglasFacturacion.PendientePagoSeguro(mo_DOCuentaAtencion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtEstadoCuenta.Text) Then
            mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar ml_idCuentaAtencion, False, 0
            MsgBox "La Cuenta se ha Cerrado correctamente", vbInformation, "Facturación"
            LimparDatos
        Else
            MsgBox "No se pudo cerrar la cuenta", vbInformation, "Facturación"
        End If
    End If
End Sub

Private Sub btnRecalculaPlan_Click()
    If mo_ReglasSISgalenhos.SisFUAyaFueEnviadoAlSisLIMA(ml_idCuentaAtencion, lnIdTipoFinanciamientoActual, wxParametro302) = True Then
       Exit Sub
    End If
    If txtPaciente.Text = "" Then
        MsgBox "Tiene que LEER el Estado de Cuenta del Paciente", vbInformation, "Resultado"
        Exit Sub
    End If
    If cmbFuenteFinanciamiento.Text = "" Then
        MsgBox "Tiene que elejir el NUEVO Plan de Atención del Paciente", vbInformation, "Resultado"
        Exit Sub
    End If
    If cmbFormaPago.Text = "" Then
        MsgBox "Tiene que elejir el NUEVO Tipo de Financiamiento del Paciente", vbInformation, "Resultado"
        Exit Sub
    End If
    If lnIdPlanActual = Val(cmbFuenteFinanciamiento.BoundText) And lnIdTipoFinanciamientoActual = Val(cmbFormaPago.BoundText) Then
        MsgBox "Tiene que elejir el NUEVO 'Plan' y 'Tipo Financiamiento'" & Chr(13) & "diferente a la ACTUAL", vbInformation, "Resultado"
        Exit Sub
    End If
    If Not (ml_idEstadoCuentaAtencion = sghAbierto Or ml_idEstadoCuentaAtencion = sghConAltaMedica) Then
        MsgBox "Verifique el Estado de la Cuenta" & Chr(13) & "deberá estar con Estado 'Abierta' o 'Alta Médica'", vbInformation, "Resultado"
        Exit Sub
    End If
    Dim oConexion As New Connection
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    '********daniel-20/12/2009-inicio***********
    
    Select Case ml_idUsuarioConPermisoEnSISoEXOoSOAT
    Case 2     'SIS
'        If mo_ReglasFacturacion.TiposFinanciamientoDevuelveComoSeTrabajaEnEstadoCuenta(Val(cmbFormaPago.BoundText), oConexion) <> sghTrabajaSeguroSIS Then
'            MsgBox "El nuevo TIPO DE FINANCIAMIENTO solo pueder ser SIS" & Chr(13) & "Chequee permiso como USUARIO en la lista: 'LABORA EN'", vbInformation, "Resultado"
'            Exit Sub
'        End If
    Case 3     'SOAT
'        If Not (Val(cmbFuenteFinanciamiento.BoundText) = 5 Or mo_ReglasFacturacion.TiposFinanciamientoDevuelveComoSeTrabajaEnEstadoCuenta(Val(cmbFormaPago.BoundText), oConexion) = sghTrabajaSeguroSOAT) Then
'            MsgBox "El nuevo TIPO DE FINANCIAMIENTO solo pueder ser PARTICULAR_HOSPITALIZADO o SOAT" & Chr(13) & "Chequee permiso como USUARIO en la lista: 'LABORA EN'", vbInformation, "Resultado"
'            Exit Sub
'        End If
    Case 4     'Convenio
'        If mo_ReglasFacturacion.TiposFinanciamientoDevuelveComoSeTrabajaEnEstadoCuenta(Val(cmbFormaPago.BoundText), oConexion) <> sghTrabajaSeguroConvenios Then
'            MsgBox "El nuevo TIPO DE FINANCIAMIENTO solo pueder ser CONVENIOS" & Chr(13) & "Chequee permiso como USUARIO en la lista: 'LABORA EN'", vbInformation, "Resultado"
'            Exit Sub
'        End If
    Case 9  'Servicio Social
    Case Else
        MsgBox "Chequee permiso como USUARIO en la lista: 'LABORA EN'", vbInformation, "Resultado"
        Exit Sub
    End Select
    '********daniel-20/12/2009-fin***********
    If MsgBox("     ¿Está seguro del Cambio       " & Chr(13) & _
              "   Plan de Cuenta de Atención ?    ", vbQuestion + vbYesNo, "Estado de Cuenta") = vbYes Then
        Dim Login As New Login
        Login.UsuarioDeEstadoDeCuenta = ml_idUsuario
        Login.CargaDesdeOtraOpcion = True
        Login.Show vbModal
        If Not Login.Autenticado Or Login.IdUsuarioAutenticado <> ml_idUsuario Then
            Exit Sub
        End If
        Dim oGrabaDatos As New SighFacturacion.dllFactUcEstadoCuenta
        oGrabaDatos.GrabaCantidadesPreciosEnElNuevoPlan Val(txtCuenta.Text), Val(cmbFuenteFinanciamiento.BoundText), Val(cmbFormaPago.BoundText), ml_idUsuario, ml_idAtencion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "Cta: " & Trim(txtCuenta.Text) & "  Nuev: " & Trim(cmbFuenteFinanciamiento.Text) & "  (Ant:" & Trim(txtEstadoCuenta.Text) & ")", lnIdTipoFinanciamientoActual, IIf(chkSoatParticular.Value = 1, True, False)
        mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar ml_idCuentaAtencion, False, 0
        MsgBox "La Cuenta paso a: " & cmbFuenteFinanciamiento.Text, vbInformation, "Facturación"
        LimparDatos
        Set oGrabaDatos = Nothing
    End If
    oConexion.Close
    Set oConexion = Nothing
End Sub

Private Sub chkFarmaciaTodos_Click()
   If ml_idUsuarioConPermisoEnSISoEXOoSOAT = 9 Then
        If chkFarmaciaTodos.Value = 1 Then
           ucFacturacionBienes.ActualizaExoneracionesPorPorcentaje True
        Else
           ucFacturacionBienes.ActualizaExoneracionesPorPorcentaje False
        End If
   Else
        If chkFarmaciaTodos.Value = 1 Then
           ucFacturacionBienes.CargaTodaLaCantidadPedidaHaciaCantidadSisSoat ml_idUsuarioConPermisoEnSISoEXOoSOAT, True
        Else
           ucFacturacionBienes.CargaTodaLaCantidadPedidaHaciaCantidadSisSoat ml_idUsuarioConPermisoEnSISoEXOoSOAT, False
        End If
   End If
End Sub

Private Sub chkServiciosTodos_Click()
   If ml_idUsuarioConPermisoEnSISoEXOoSOAT = 9 Then
        If chkServiciosTodos.Value = 1 Then
           ucFacturacionServicios.ActualizaExoneracionesPorPorcentaje True
        Else
           ucFacturacionServicios.ActualizaExoneracionesPorPorcentaje False
        End If
   Else
        If chkServiciosTodos.Value = 1 Then
           ucFacturacionServicios.CargaTodaLaCantidadPedidaHaciaCantidadSisSoat ml_idUsuarioConPermisoEnSISoEXOoSOAT, True
        Else
           ucFacturacionServicios.CargaTodaLaCantidadPedidaHaciaCantidadSisSoat ml_idUsuarioConPermisoEnSISoEXOoSOAT, False
        End If
   End If
End Sub

Private Sub cmbAgrupar_Click()
    
    If cmbAgrupar.Text <> "<Sin agrupar>" And txtNroHistoria.Text <> "" Then
        ObtenerDatosDeLaCuentaDeAtencion
    End If
    
End Sub




Private Sub cmbFuenteFinanciamiento_Change()
    If lnIdTipoFinanciamientoActual = sghTipoFinanciamiento.sghSOAT And cmbFuenteFinanciamiento.BoundText = "5" Then
       chkSoatParticular.Visible = True
    Else
       chkSoatParticular.Visible = False
    End If

End Sub

Private Sub cmbFuenteFinanciamiento_Click(Area As Integer)
        cmbFuenteFinanciamiento_Change
        '
        Set oRsFormaPago = mo_ReglasFacturacion.TiposFinanciamientosTarifaSeleccionarPorPlan(Val(cmbFuenteFinanciamiento.BoundText))
        Set cmbFormaPago.RowSource = oRsFormaPago
        cmbFormaPago.ListField = "Descripcion"
        cmbFormaPago.BoundColumn = "idTipoFinanciamiento"
        mo_Formulario.HabilitarDeshabilitar cmbFormaPago, True
        If oRsFormaPago.RecordCount = 1 Then
           cmbFormaPago.BoundText = oRsFormaPago.Fields!idTipoFinanciamiento
        ElseIf Val(cmbFuenteFinanciamiento.BoundText) = 5 Then
           cmbFormaPago.BoundText = "1"
        Else
           cmbFormaPago.Text = ""
        End If

End Sub

Private Sub cmdBuscaCuentaPorApellidos_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaPacientes
    Dim oDOPaciente As New doPaciente
    Dim oConexion As New Connection
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oBusqueda.TipoFiltro = sghFiltrarTodos
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oConexion)
        If Not oDOPaciente Is Nothing Then
           mo_cmbIdTipoGenHistoriaClinica.BoundText = oDOPaciente.IdTipoNumeracion
           txtNroHistoria.Text = oDOPaciente.NroHistoriaClinica
           txtNroOrdenPagoS.Text = ""
           txtNroHistoria_KeyPress 13
           
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oDOPaciente = Nothing
    Set oBusqueda = Nothing
End Sub


Private Sub cmdImprimeCtaPorServicioHosp_Click()
'    Dim oExcel As Excel.Application
'    Dim oWorkBookPlantilla As Workbook
'    Dim oWorkBook As Workbook
'    Dim oWorkSheet As Worksheet
'    Dim iFila As Long: Dim iCol As Integer
'    Dim rsReporte As New Recordset
'    Dim ms_EstadosFacturacion As String
'    Dim ms_TiposFinanciamiento As String
'    Dim ml_AgruparPor As Long
'    Dim mo_ReporteUtil As New ReporteUtil
'    Dim idPuntoCarga As Long
'
'    Dim lnTSubTotal As Double: Dim lnTSubTotalSIS As Double: Dim lnTSubTotalSOAT As Double: Dim lnTSubTotalEXO As Double: Dim lnTsubTotalConvenio As Double
'    Dim lnTSubTotalPAGO As Double: Dim lnTSubTotalDEBE As Double: Dim lnTSubTotalSALDO As Double
'
'    Dim lnTotal As Double: Dim lnTotalSIS As Double: Dim lnTotalSOAT As Double: Dim lnTotalEXO As Double: Dim lnTotalConvenio As Double
'    Dim lnTotalPAGO As Double: Dim lnTotalDEBE As Double: Dim lnTotalSALDO As Double
'
'    Dim lnDebe As Double: Dim lnPago As Double: Dim lnSaldo As Double
'    Dim lnTDebe As Double: Dim lnTPago As Double: Dim lnTSaldo As Double
'    Dim lnSIS As Double: Dim lnSOAT As Double: Dim lnEXO As Double: Dim lnTotalCredito As Double: Dim lnConvenio As Double
'    Dim lnDctos As Double: Dim lnPagoEnFarmacia As Double: Dim lnPagoEnServicios As Double
'    Dim lnTotalPagosAdelantados As Double
'    Dim lcBuscaParametro As New SIGHDatos.Parametros
'    Dim lnCantidadPagarBienes As Long, lnTotalPagarBienes As Double
'    Dim ldFechaAlta As Date, lcHoraAlta As String
'    Dim lnFor As Integer, lnForNum As Integer
'    Dim oGenerarRecordsetProductos As New SighFacturacion.dllFactUcEstadoCuenta
'    Dim lnTotalPagarEstancia As Double, lnTotalDiasEstancia As Long
'    Dim oConexion As New Connection
'    oConexion.Open SIGHEntidades.CadenaConexion
'    oConexion.CursorLocation = adUseClient
'    If txtPaciente.Text = "" Then
'        MsgBox "Tiene que LEER el Estado de Cuenta del Paciente", vbInformation, "Resultado"
'    Else
'        If ucFacturacionBienes.FacturacionProductos.RecordCount = 0 And ucFacturacionServicios.FacturacionProductos.RecordCount = 0 Then
'           MsgBox "Tiene que LEER el Estado de Cuenta del Paciente", vbInformation, "Resultado"
'           Exit Sub
'        End If
'        MousePointer = 11
'        'Crea nueva hoja
'        Set oExcel = GalenhosExcelApplication()  'New Excel.Application
'        Set oWorkBook = oExcel.Workbooks.Add
'        'Abre, copia y cierra la plantilla
'        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\ECuentaCte.xls")
'        oWorkBookPlantilla.Worksheets("CuentaCte").Copy Before:=oWorkBook.Sheets(1)
'        oWorkBookPlantilla.Close
'        'Activa la primera hoja
'        Set oWorkSheet = oWorkBook.Sheets(1)
'        'oWorkSheet.PageSetup.LeftHeader = lcBuscaParametro.SeleccionaFilaParametro(205)
'        oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\Imagen de reportes.jpg"
'        'Inicio de Impresion
'
'        If ml_IdTipoServicio = sghConsultaExterna Then
'           Set rsReporte = mo_ReglasFacturacion.AtencionesFiltraIngresosPorIdAtencion(ml_idAtencion)
'        Else
'           Set rsReporte = mo_ReglasFacturacion.AtencionesFiltraEgresosPorIdAtencion(ml_idAtencion)
'        End If
'        If rsReporte.RecordCount = 0 Then
'           MousePointer = 1
'           Exit Sub
'        End If
'        oWorkSheet.Cells(1, 3).Value = "Estado de Cuenta al " & lcBuscaParametro.RetornaFechaServidorSQL()
'        oWorkSheet.Cells(3, 3).Value = Trim(Str(ml_idCuentaAtencion)) & "  " & txtEstadoCuenta.Text
'        oWorkSheet.Cells(3, 8).Value = Trim(Str(ml_idAtencion)) & IIf(txtDxEgr.Text = "", "", "      Dx Egreso: " & txtDxEgr.Text)
'        oWorkSheet.Cells(4, 3).Value = txtPaciente.Text
'        oWorkSheet.Cells(4, 8).Value = Trim(txtNroHistoria.Text) & "       Dom.Pac: " & Trim(txtDomicilioPacienteEnAtencion.Text)
'        oWorkSheet.Cells(5, 3).Value = "'" & Format(rsReporte.Fields!FechaIngreso, SIGHEntidades.DevuelveFechaSoloFormato_DMY) & " " & rsReporte.Fields!HoraIngreso
'        oWorkSheet.Cells(5, 8).Value = IIf(IsNull(rsReporte.Fields!codServicio), "", rsReporte.Fields!codServicio & " - " & rsReporte.Fields!DServicio)
'        oWorkSheet.Cells(6, 3).Value = IIf(IsNull(rsReporte.Fields!FechaEgreso), "", "'" & Format(rsReporte.Fields!FechaEgreso & " " & rsReporte.Fields!HoraEgreso, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HM))
'        oWorkSheet.Cells(6, 8).Value = IIf(IsNull(rsReporte.Fields!codCama), "", rsReporte.Fields!codCama)
'        iFila = 8
'        iCol = 2
'        ms_EstadosFacturacion = ""
'        ms_TiposFinanciamiento = ""
'        ml_AgruparPor = 1
'        lnTotal = 0: lnTotalSIS = 0: lnTotalSOAT = 0: lnTotalEXO = 0: lnTotalPAGO = 0: lnTotalDEBE = 0: lnTotalSALDO = 0: lnTotalCredito = 0
'
'        lnTDebe = 0: lnTPago = 0: lnTSaldo = 0
'
'        'Farmacia
'        Set rsReporte = ucFacturacionBienes.FacturacionProductos
'        If rsReporte.RecordCount > 0 Then
'            rsReporte.Sort = "idServicioDeEstancia"
'            rsReporte.MoveFirst
'            Do While Not rsReporte.EOF
'                lnTSubTotal = 0: lnTSubTotalSIS = 0: lnTSubTotalSOAT = 0: lnTSubTotalEXO = 0: lnTsubTotalConvenio = 0
'                lnTSubTotalPAGO = 0: lnTSubTotalDEBE = 0: lnTSubTotalSALDO = 0
'                idPuntoCarga = rsReporte.Fields("idServicioDeEstancia").Value
'                oWorkSheet.Cells(iFila, 2).Value = "(Farmacia) " & rsReporte.Fields!ServicioDeEstancia
'                iFila = iFila + 1
'                Do While Not rsReporte.EOF And idPuntoCarga = rsReporte.Fields("idServicioDeEstancia").Value
'                   If rsReporte.Fields!IdEstadoFacturacion = 4 Or rsReporte.Fields!IdEstadoFacturacion = 1 Then   'Solo PAGADOS y REGISTRADOS
'                        lnSIS = IIf(IsNull(rsReporte.Fields!ImporteSIS), 0, rsReporte.Fields!ImporteSIS)
'                        lnEXO = IIf(IsNull(rsReporte.Fields!importeEXO), 0, rsReporte.Fields!importeEXO)
'                        lnSOAT = IIf(IsNull(rsReporte.Fields!ImporteSOAT), 0, rsReporte.Fields!ImporteSOAT)
'                        lnConvenio = IIf(IsNull(rsReporte.Fields!ImporteConv), 0, rsReporte.Fields!ImporteConv)
'                        oWorkSheet.Cells(iFila, 3).Value = rsReporte.Fields("Codigo").Value & " - " & rsReporte.Fields("NombreProducto").Value
'                        oWorkSheet.Cells(iFila, 5).Value = rsReporte.Fields("CantidadPagar").Value
'                        oWorkSheet.Cells(iFila, 6).Value = rsReporte.Fields("preciounitario").Value
'                        oWorkSheet.Cells(iFila, 7).Value = Format(rsReporte.Fields("TotalPagar").Value, "####,###.#0")
'                        oWorkSheet.Cells(iFila, 8).Value = Format(lnEXO, "####,###.#0")
'                        oWorkSheet.Cells(iFila, 9).Value = Format(lnSIS, "####,###.#0")
'                        oWorkSheet.Cells(iFila, 10).Value = Format(lnSOAT, "####,###.#0")
'                        oWorkSheet.Cells(iFila, 11).Value = Format(lnConvenio, "####,###.#0")
'
'                        If lbGeneraReciboPago = True Then
'                           lnDebe = rsReporte.Fields!TotalPagar - lnEXO - lnSIS - lnSOAT
'                        Else
'                           If (rsReporte.Fields!CantidadSIS + rsReporte.Fields!CantidadSOAT + rsReporte.Fields!cantidadConv) = 0 Then
'                               lnDebe = rsReporte.Fields!TotalPagar
'                           Else
'                               lnDebe = 0
'                           End If
'                        End If
'                        If rsReporte.Fields!IdEstadoFacturacion = 4 Then
'                           lnPago = rsReporte.Fields("ImporteEnBoleta").Value - lnEXO
'                        Else
'                           lnPago = 0
'                        End If
'                        lnSaldo = lnDebe - lnPago
'                        oWorkSheet.Cells(iFila, 12).Value = Format(lnDebe, "####,###.#0")
'                        oWorkSheet.Cells(iFila, 13).Value = Format(lnPago, "####,###.#0")
'                        oWorkSheet.Cells(iFila, 14).Value = Format(lnSaldo, "####,###.#0")
'                        oWorkSheet.Cells(iFila, 15).Value = rsReporte.Fields!nroDcto
'                        oWorkSheet.Cells(iFila, 16).Value = rsReporte.Fields!FechaDespacho
'
'                        lnTSubTotal = lnTSubTotal + rsReporte.Fields("TotalPagar").Value
'                        lnTSubTotalEXO = lnTSubTotalEXO + lnEXO
'                        lnTSubTotalSIS = lnTSubTotalSIS + lnSIS
'                        lnTSubTotalSOAT = lnTSubTotalSOAT + lnSOAT
'                        lnTsubTotalConvenio = lnTsubTotalConvenio + lnConvenio
'                        lnTSubTotalPAGO = lnTSubTotalPAGO + lnPago
'                        lnTSubTotalDEBE = lnTSubTotalDEBE + lnDebe
'                        lnTSubTotalSALDO = lnTSubTotalSALDO + lnSaldo
'
'                        lnTotal = lnTotal + rsReporte.Fields("TotalPagar").Value
'                        lnTotalSIS = lnTotalSIS + lnSIS
'                        lnTotalSOAT = lnTotalSOAT + lnSOAT
'                        lnTotalEXO = lnTotalEXO + lnEXO
'                        lnTotalConvenio = lnTotalConvenio + lnConvenio
'                        lnTotalPAGO = lnTotalPAGO + lnPago
'                        lnTotalDEBE = lnTotalDEBE + lnDebe
'                        lnTotalSALDO = lnTotalSALDO + lnSaldo
'
'                        If rsReporte.Fields!idProducto = lnIdPagosACuenta Then   'Pagos a cuenta
'                           lnTotalCredito = lnTotalCredito + rsReporte.Fields("TotalPorPagar").Value
'                        End If
'
'                        iFila = iFila + 1
'                    End If
'                    rsReporte.MoveNext
'                    If rsReporte.EOF Then
'                       Exit Do
'                    End If
'                 Loop
'                 mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 3, iFila, iCol + 14
'                 oWorkSheet.Cells(iFila, 7).Value = Format(lnTSubTotal, "####,###.#0")
'                 oWorkSheet.Cells(iFila, 8).Value = Format(lnTSubTotalEXO, "####,###.#0")
'                 oWorkSheet.Cells(iFila, 9).Value = Format(lnTSubTotalSIS, "####,###.#0")
'                 oWorkSheet.Cells(iFila, 10).Value = Format(lnTSubTotalSOAT, "####,###.#0")
'                 oWorkSheet.Cells(iFila, 11).Value = Format(lnTsubTotalConvenio, "####,###.#0")
'                 oWorkSheet.Cells(iFila, 12).Value = Format(lnTSubTotalDEBE, "####,###.#0")
'                 oWorkSheet.Cells(iFila, 13).Value = Format(lnTSubTotalPAGO, "####,###.#0")
'                 oWorkSheet.Cells(iFila, 14).Value = Format(lnTSubTotalSALDO, "####,###.#0")
'                 iFila = iFila + 1
'             Loop
'        End If
'        rsReporte.Close
'        lnPagoEnFarmacia = lnTSubTotalSALDO
'        'Servicios
'        lnTotalPagarEstancia = 0
'        lnFor = 1
'        If ml_IdTipoServicio = sghHospitalizacion Then
'           lnFor = 2
'        End If
'        For lnForNum = 1 To lnFor
'            If lnForNum = 1 Then
'                Set rsReporte = ucFacturacionServicios.FacturacionProductos
'                If ml_IdTipoServicio = sghHospitalizacion Then
'                   On Error Resume Next
'                   rsReporte.Filter = "idPuntoCarga<>" & sghPtoCargaAdmisionHospitalizacion
'                End If
'            Else
'                rsReporte.Filter = ""
'                If txtFegreso.Text = "" Then
'                    ldFechaAlta = CDate(Format(Now, SIGHEntidades.DevuelveFechaSoloFormato_DMY))
'                    lcHoraAlta = Format(Now, SIGHEntidades.DevuelveHoraSoloFormato_HM)
'                Else
'                    ldFechaAlta = CDate(Format(CDate(txtFegreso.Text), SIGHEntidades.DevuelveFechaSoloFormato_DMY))
'                    lcHoraAlta = Format(CDate(txtFegreso.Text), SIGHEntidades.DevuelveHoraSoloFormato_HM)
'                End If
'                oGenerarRecordsetProductos.GenerarRecordsetProductos rsReporte
'                mo_AdminAdmision.GeneraEstanciaPorCadaServicioHospitalizado ml_idCuentaAtencion, ldFechaAlta, lcHoraAlta, rsReporte, lnTotalPagarEstancia, lnTotalDiasEstancia, oConexion
'                If txtFegreso.Text <> "" Then
'                   lnTotalPagarEstancia = 0
'                End If
'            End If
'            'Set rsReporte = ucFacturacionServicios.FacturacionProductos
'            lnTotalPagosAdelantados = 0
'            If rsReporte.RecordCount > 0 Then
'                rsReporte.Sort = "idServicioDeEstancia"
'                rsReporte.MoveFirst
'                Do While Not rsReporte.EOF
'                    idPuntoCarga = rsReporte.Fields("idServicioDeEstancia").Value
'                    oWorkSheet.Cells(iFila, 2).Value = "(Servicios) " & rsReporte.Fields!ServicioDeEstancia
'                    lnTSubTotal = 0: lnTSubTotalSIS = 0: lnTSubTotalSOAT = 0: lnTSubTotalEXO = 0: lnTsubTotalConvenio = 0
'                    lnTSubTotalPAGO = 0: lnTSubTotalDEBE = 0: lnTSubTotalSALDO = 0
'                    iFila = iFila + 1
'                    Do While Not rsReporte.EOF And idPuntoCarga = rsReporte.Fields("idServicioDeEstancia").Value
'                        If rsReporte.Fields!IdEstadoFacturacion = 4 Or rsReporte.Fields!IdEstadoFacturacion = 1 Or rsReporte.Fields!IdEstadoFacturacion = 10 Or rsReporte.Fields!IdEstadoFacturacion = sghConPreVenta Then  'Solo PAGADOS/REGISTRADOS/AUTORIZ.AUTOMATICA/preventa
'                            lnSIS = IIf(IsNull(rsReporte.Fields!ImporteSIS), 0, rsReporte.Fields!ImporteSIS)
'                            lnSOAT = IIf(IsNull(rsReporte.Fields!ImporteSOAT), 0, rsReporte.Fields!ImporteSOAT)
'                            lnEXO = IIf(IsNull(rsReporte.Fields!importeEXO), 0, rsReporte.Fields!importeEXO)
'                            lnConvenio = IIf(IsNull(rsReporte.Fields!ImporteConv), 0, rsReporte.Fields!ImporteConv)
'                            oWorkSheet.Cells(iFila, 3).Value = rsReporte.Fields("Codigo").Value & " - " & rsReporte.Fields("NombreProducto").Value
'                            oWorkSheet.Cells(iFila, 5).Value = rsReporte.Fields("CantidadPagar").Value
'                            oWorkSheet.Cells(iFila, 6).Value = rsReporte.Fields("preciounitario").Value
'                            oWorkSheet.Cells(iFila, 7).Value = Format(rsReporte.Fields("TotalPagar").Value, "####,###.#0")
'                            oWorkSheet.Cells(iFila, 8).Value = Format(lnEXO, "####,###.#0")
'                            oWorkSheet.Cells(iFila, 9).Value = Format(lnSIS, "####,###.#0")
'                            oWorkSheet.Cells(iFila, 10).Value = Format(lnSOAT, "####,###.#0")
'                            oWorkSheet.Cells(iFila, 11).Value = Format(lnConvenio, "####,###.#0")
'                            If rsReporte.Fields!IdEstadoFacturacion = 4 Then
'                                lnPago = rsReporte.Fields("ImporteEnBoleta").Value - lnEXO
'                            Else
'                                lnPago = 0
'                            End If
'                            If rsReporte.Fields!idProducto <> lnIdPagosACuenta And rsReporte.Fields!idProducto <> lnIdPagosXdevoluciones Then
'                               If lbGeneraReciboPago = True Then
'                                    lnDebe = rsReporte.Fields!TotalPagar - lnEXO - lnSIS - lnSOAT
'                               Else
'                                    If (rsReporte.Fields!CantidadSIS + rsReporte.Fields!CantidadSOAT + rsReporte.Fields!cantidadConv) = 0 Then
'                                        lnDebe = rsReporte.Fields!TotalPagar
'                                    Else
'                                        lnDebe = 0
'                                    End If
'                               End If
'                               lnSaldo = lnDebe - lnPago
'                            Else
'                               If rsReporte.Fields!idProducto = lnIdPagosACuenta Then
'                                    lnTotalPagosAdelantados = lnTotalPagosAdelantados + rsReporte.Fields!ImporteEnBoleta
'                                    lnDebe = 0
'                                    If lbGeneraReciboPago = True Then
'                                       lnSaldo = -rsReporte.Fields!ImporteEnBoleta
'                                    Else
'                                       lnSaldo = 0
'                                    End If
'                               Else
'                                    lnDebe = 0
'                                    If lbGeneraReciboPago = True Then
'                                       lnSaldo = rsReporte.Fields!ImporteEnBoleta
'                                       lnPago = -rsReporte.Fields!ImporteEnBoleta
'                                    Else
'                                       lnSaldo = 0
'                                       lnPago = -rsReporte.Fields!ImporteEnBoleta
'                                    End If
'                               End If
'                            End If
'
'                            oWorkSheet.Cells(iFila, 12).Value = Format(lnDebe, "####,###.#0")
'                            oWorkSheet.Cells(iFila, 13).Value = Format(lnPago, "####,###.#0")
'                            oWorkSheet.Cells(iFila, 14).Value = Format(lnSaldo, "####,###.#0")
'                            oWorkSheet.Cells(iFila, 15).Value = rsReporte.Fields!nroDcto
'                            oWorkSheet.Cells(iFila, 16).Value = rsReporte.Fields!FechaDespacho
'
'                            If rsReporte.Fields!idProducto <> lnIdPagosACuenta Then
'                               lnTSubTotal = lnTSubTotal + rsReporte.Fields("TotalPagar").Value
'                               lnTSubTotalDEBE = lnTSubTotalDEBE + lnDebe
'                            End If
'                            lnTSubTotalEXO = lnTSubTotalEXO + lnEXO
'                            lnTSubTotalSIS = lnTSubTotalSIS + lnSIS
'                            lnTSubTotalSOAT = lnTSubTotalSOAT + lnSOAT
'                            lnTsubTotalConvenio = lnTsubTotalConvenio + lnConvenio
'                            lnTSubTotalPAGO = lnTSubTotalPAGO + lnPago
'                            lnTSubTotalSALDO = lnTSubTotalSALDO + lnSaldo
'
'                            If rsReporte.Fields!idProducto <> lnIdPagosACuenta Then
'                               lnTotal = lnTotal + rsReporte.Fields("TotalPagar").Value
'                               lnTotalDEBE = lnTotalDEBE + lnDebe
'                            End If
'                            lnTotalSIS = lnTotalSIS + lnSIS
'                            lnTotalSOAT = lnTotalSOAT + lnSOAT
'                            lnTotalEXO = lnTotalEXO + lnEXO
'                            lnTotalConvenio = lnTotalConvenio + lnConvenio
'                            lnTotalPAGO = lnTotalPAGO + lnPago
'                            lnTotalSALDO = lnTotalSALDO + lnSaldo
'
'                            If rsReporte.Fields!idProducto = lnIdPagosACuenta Then   'Pagos a cuenta
'                               lnTotalCredito = lnTotalCredito + rsReporte.Fields("TotalPorPagar").Value
'                            End If
'
'                            iFila = iFila + 1
'                        End If
'                        rsReporte.MoveNext
'                        If rsReporte.EOF Then
'                           Exit Do
'                        End If
'                    Loop
'                    mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 3, iFila, iCol + 14
'                    oWorkSheet.Cells(iFila, 7).Value = Format(lnTSubTotal, "####,###.#0")
'                    oWorkSheet.Cells(iFila, 8).Value = Format(lnTSubTotalEXO, "####,###.#0")
'                    oWorkSheet.Cells(iFila, 9).Value = Format(lnTSubTotalSIS, "####,###.#0")
'                    oWorkSheet.Cells(iFila, 10).Value = Format(lnTSubTotalSOAT, "####,###.#0")
'                    oWorkSheet.Cells(iFila, 11).Value = Format(lnTsubTotalConvenio, "####,###.#0")
'                    oWorkSheet.Cells(iFila, 12).Value = Format(lnTSubTotalDEBE, "####,###.#0")
'                    oWorkSheet.Cells(iFila, 13).Value = Format(lnTSubTotalPAGO, "####,###.#0")
'                    oWorkSheet.Cells(iFila, 14).Value = Format(lnTSubTotalSALDO, "####,###.#0")
'                    iFila = iFila + 1
'                 Loop
'            End If
'            iFila = iFila + 1
'        Next
'        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, iCol + 14
'        oWorkSheet.Cells(iFila, 2).Value = "Total: "
'        oWorkSheet.Cells(iFila, 7).Value = Format(lnTotal, "####,###.#0")
'        oWorkSheet.Cells(iFila, 8).Value = Format(lnTotalEXO, "####,###.#0")
'        oWorkSheet.Cells(iFila, 9).Value = Format(lnTotalSIS, "####,###.#0")
'        oWorkSheet.Cells(iFila, 10).Value = Format(lnTotalSOAT, "####,###.#0")
'        oWorkSheet.Cells(iFila, 11).Value = Format(lnTotalConvenio, "####,###.#0")
'        oWorkSheet.Cells(iFila, 12).Value = Format(lnTotalDEBE, "####,###.#0")
'        oWorkSheet.Cells(iFila, 13).Value = Format(lnTotalPAGO, "####,###.#0")
'        oWorkSheet.Cells(iFila, 14).Value = Format(lnTotalSALDO, "####,###.#0")
'        If lbGeneraReciboPago = True Then
'            If lnTotalPagosAdelantados > lnPagoEnFarmacia Then
'               lnTotalPagosAdelantados = lnTotalPagosAdelantados - lnPagoEnFarmacia
'               lnPagoEnFarmacia = 0
'            Else
'               lnPagoEnFarmacia = lnPagoEnFarmacia - lnTotalPagosAdelantados
'               lnTotalPagosAdelantados = 0
'            End If
'            lnPagoEnServicios = lnPagoEnServicios - lnTotalPagosAdelantados
'        Else
'           Select Case mo_ReglasFacturacion.TiposFinanciamientoDevuelveComoSeTrabajaEnEstadoCuenta(lnIdTipoFinanciamientoActual, oConexion)
'           Case sghTrabajaSeguroSIS
'               lnTotalSIS = lnTotalSIS - lnTotalPagosAdelantados
'           Case sghTrabajaSeguroSOAT
'               lnTotalSOAT = lnTotalSOAT - lnTotalPagosAdelantados
'           Case sghTrabajaSeguroConvenios
'               lnTotalConvenio = lnTotalConvenio - lnTotalPagosAdelantados
'           End Select
'        End If
'        lnPagoEnServicios = lnTotalSALDO - lnPagoEnFarmacia
'
'        iFila = iFila + 3
'        oWorkSheet.Cells(iFila, 11).Value = "TOTAL CUENTA"
'        oWorkSheet.Cells(iFila, 12).Value = Format(lnTotal, "####,###.#0")
'        iFila = iFila + 1
'        oWorkSheet.Cells(iFila, 11).Value = "EXONERADO"
'        oWorkSheet.Cells(iFila, 12).Value = Format(lnTotalEXO, "####,###.#0")
'        iFila = iFila + 1
'        oWorkSheet.Cells(iFila, 11).Value = "SIS CUBRE (-PAGOS A CUENTA)"
'        oWorkSheet.Cells(iFila, 12).Value = Format(lnTotalSIS, "####,###.#0")
'        iFila = iFila + 1
'        oWorkSheet.Cells(iFila, 11).Value = "SOAT CUBRE (-PAGOS A CUENTA)"
'        oWorkSheet.Cells(iFila, 12).Value = Format(lnTotalSOAT, "####,###.#0")
'        iFila = iFila + 1
'        oWorkSheet.Cells(iFila, 11).Value = "CONVENIOS CUBRE (-PAGOS A CUENTA)"
'        oWorkSheet.Cells(iFila, 12).Value = Format(lnTotalConvenio, "####,###.#0")
'        iFila = iFila + 1
'        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 10, iFila, iCol + 10
'        oWorkSheet.Cells(iFila, 11).Value = "TOTAL DEUDA"
'        oWorkSheet.Cells(iFila, 12).Value = Format(lnTotalDEBE, "####,###.#0")
'        iFila = iFila + 1
'        oWorkSheet.Cells(iFila, 11).Value = "PAGOS REALIZADOS"
'        oWorkSheet.Cells(iFila, 12).Value = Format(lnTotalPAGO, "####,###.#0")
'        iFila = iFila + 1
'        oWorkSheet.Cells(iFila, 11).Value = "CAJA DEBE INGRESAR"
'        iFila = iFila + 1
'        oWorkSheet.Cells(iFila, 11).Value = "CREDITO"
'        oWorkSheet.Cells(iFila, 12).Value = Format(lnTotalCredito, "####,###.#0")
'        iFila = iFila + 1
'        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 10, iFila, iCol + 10
'
'        lnPagoEnServicios = CCur(txtTotalServicios.Text) + lnTotalPagarEstancia
'        lnPagoEnFarmacia = CCur(txtTotalFarmacia.Text)
'        lnTotalSALDO = lnPagoEnServicios + lnPagoEnFarmacia
'
'        oWorkSheet.Cells(iFila, 11).Value = "PACIENTE DEBE PAGAR"
'        oWorkSheet.Cells(iFila, 12).Value = Format(lnTotalSALDO, "####,###.#0")
'        iFila = iFila + 1
'        oWorkSheet.Cells(iFila, 11).Value = "PAGO POR CONSUMO FARMACIA"
'        oWorkSheet.Cells(iFila, 12).Value = Format(lnPagoEnFarmacia, "####,###.#0")
'        iFila = iFila + 1
'        oWorkSheet.Cells(iFila, 11).Value = "PAGO POR CONSUMO SERVICIO"
'        oWorkSheet.Cells(iFila, 12).Value = Format(lnPagoEnServicios, "####,###.#0")
'
'        'Transferencias
'        rsReporte.Close
'        Set rsReporte = mo_ReglasFacturacion.AtencionesFiltraTransferenciasPorIdAtencion(ml_idAtencion)
'        If rsReporte.RecordCount > 0 Then
'            iFila = iFila - rsReporte.RecordCount - 1
'            oWorkSheet.Cells(iFila, 2).Value = "ESTADIA"
'            iFila = iFila + 1
'            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 6
'            oWorkSheet.Cells(iFila, 2).Value = "Cod.Cama"
'            oWorkSheet.Cells(iFila, 3).Value = "Servicio que fue transferido"
'            oWorkSheet.Cells(iFila, 5).Value = "F.Transf"
'            oWorkSheet.Cells(iFila, 6).Value = "H.Transf"
'            iFila = iFila + 1
'            rsReporte.MoveFirst
'            Do While Not rsReporte.EOF
'                oWorkSheet.Cells(iFila, 2).Value = rsReporte!CodigoCama
'                oWorkSheet.Cells(iFila, 3).Value = rsReporte!NombreServicio
'                oWorkSheet.Cells(iFila, 5).Value = "'" & rsReporte!FechaOcupacion
'                oWorkSheet.Cells(iFila, 6).Value = rsReporte!HoraOcupacion
'                'oWorkSheet.Cells(iFila, 7).Value = rsReporte!NombreMedico
'                iFila = iFila + 1
'                rsReporte.MoveNext
'            Loop
'        End If
'        '***Donaciones en Farmacia
'        If rsItemsDonaciones.RecordCount > 0 Then
'           Dim lnCantidadDona As Long, lcCodigoDona As String, lbContinua As Boolean
'           iFila = iFila + 2
'           oWorkSheet.Cells(iFila, 2).Value = "LISTA DE DONACIONES:"
'           iFila = iFila + 1
'           oWorkSheet.Cells(iFila, 2).Value = "Descripción"
'           oWorkSheet.Cells(iFila, 9).Value = "Cantidad"
'           mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 9
'           Set rsReporte = Nothing
'           With rsReporte
'                  .Fields.Append "Codigo", adVarChar, 10, adFldIsNullable
'                  .Fields.Append "Nombre", adVarChar, 200, adFldIsNullable
'                  .Fields.Append "Cantidad", adInteger
'                  .CursorType = adOpenDynamic
'                  .LockType = adLockOptimistic
'                  .Open
'           End With
'           rsItemsDonaciones.MoveFirst
'           Do While Not rsItemsDonaciones.EOF
'              lbContinua = True
'              If rsReporte.RecordCount > 0 Then
'                 rsReporte.MoveFirst
'                 rsReporte.Find "Codigo='" & rsItemsDonaciones.Fields!Codigo & "'"
'                 If Not rsReporte.EOF Then
'                    lbContinua = False
'                 End If
'              End If
'              If lbContinua = True Then
'                  rsReporte.AddNew
'                  rsReporte.Fields!Codigo = rsItemsDonaciones.Fields!Codigo
'                  rsReporte.Fields!Nombre = rsItemsDonaciones.Fields!Nombre
'              End If
'              rsReporte.Fields!Cantidad = rsReporte.Fields!Cantidad + rsItemsDonaciones.Fields!Cantidad
'              rsReporte.Update
'              rsItemsDonaciones.MoveNext
'           Loop
'           rsReporte.Sort = "nombre,codigo"
'           rsReporte.MoveFirst
'           Do While Not rsReporte.EOF
'              iFila = iFila + 1
'              lcCodigoDona = rsReporte.Fields!Codigo
'              oWorkSheet.Cells(iFila, 2).Value = Trim(rsReporte.Fields!Codigo) & " " & rsReporte.Fields!Nombre
'              oWorkSheet.Cells(iFila, 9).Value = rsReporte.Fields!Cantidad
'              rsReporte.MoveNext
'           Loop
'        End If
'        '
'        If lcListaDeOrdenesDePago <> "" Then
'            iFila = iFila + 2
'            oWorkSheet.Cells(iFila, 2).Value = "* El CAJERO debe emitir Boletas usando " & lcListaDeOrdenesDePago
'        End If
'        '
'        oWorkSheet.PageSetup.PrintTitleRows = "$1:$7"
'        If oWorkSheet.PageSetup.PrintArea <> "" Then oWorkSheet.PageSetup.PrintArea = "$A$1:$R$" & (iFila + 2) 'sighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
'        oExcel.Visible = True
'        oWorkSheet.PrintPreview
'        'oWorkSheet.PrintOut
'    End If
'    oConexion.Close
'    Set oConexion = Nothing
'    Set oWorkSheet = Nothing
'    Set oExcel = Nothing
'    MousePointer = 1
    Dim iFila As Long: Dim iCol As Integer
    Dim rsReporte As New Recordset
    Dim ms_EstadosFacturacion As String
    Dim ms_TiposFinanciamiento As String
    Dim ml_AgruparPor As Long
    Dim mo_ReporteUtil As New sighentidades.ReporteUtil
    Dim idPuntoCarga As Long
    
    Dim lnTSubTotal As Double: Dim lnTSubTotalSIS As Double: Dim lnTSubTotalSOAT As Double: Dim lnTSubTotalEXO As Double: Dim lnTsubTotalConvenio As Double
    Dim lnTSubTotalPAGO As Double: Dim lnTSubTotalDEBE As Double: Dim lnTSubTotalSALDO As Double
    
    Dim lnTotal As Double: Dim lnTotalSIS As Double: Dim lnTotalSOAT As Double: Dim lnTotalEXO As Double: Dim lnTotalConvenio As Double
    Dim lnTotalPAGO As Double: Dim lnTotalDEBE As Double: Dim lnTotalSALDO As Double
    
    Dim lnDebe As Double: Dim lnPago As Double: Dim lnSaldo As Double
    Dim lnTDebe As Double: Dim lnTPago As Double: Dim lnTSaldo As Double
    Dim lnSIS As Double: Dim lnSOAT As Double: Dim lnEXO As Double: Dim lnTotalCredito As Double: Dim lnConvenio As Double
    Dim lnDctos As Double: Dim lnPagoEnFarmacia As Double: Dim lnPagoEnServicios As Double
    Dim lnTotalPagosAdelantados As Double
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim lnCantidadPagarBienes As Long, lnTotalPagarBienes As Double
    Dim ldFechaAlta As Date, lcHoraAlta As String
    Dim lnFor As Integer, lnForNum As Integer
    Dim oGenerarRecordsetProductos As New SighFacturacion.dllFactUcEstadoCuenta
    Dim lnTotalPagarEstancia As Double, lnTotalDiasEstancia As Long
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    
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
            Dim lnHWnd As Long
        Else
            Dim oExcel As Excel.Application
            Dim oWorkBookPlantilla As Workbook
            Dim oWorkBook As Workbook
            Dim oWorkSheet As Worksheet
            Dim oRange As range
            Dim range As Excel.range
            Dim borders As Excel.borders
        End If
    
    If txtPaciente.Text = "" Then
        MsgBox "Tiene que LEER el Estado de Cuenta del Paciente", vbInformation, "Resultado"
    Else
        If ucFacturacionBienes.FacturacionProductos.RecordCount = 0 And ucFacturacionServicios.FacturacionProductos.RecordCount = 0 Then
           MsgBox "Tiene que LEER el Estado de Cuenta del Paciente", vbInformation, "Resultado"
           Exit Sub
        End If
        MousePointer = 11
        'Crea nueva hoja
        If lbEsOpenOffice = True Then
            'Abre el archivo ExcelOpenOffice
            lcArchivoExcel = App.Path + "\Plantillas\ECuentaCte.ods"
    '        FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
    '        Chemin = "file:///" & App.Path & "\Plantillas\"
    '        Chemin = Replace(Chemin, "\", "/")
    '        Fichier = Chemin & "/OpenOffice.ods"
            '
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
            'Encabezado de Pagina
            mo_CabeceraReportes.CabeceraReportes Document, True
            ' Pone la ventana en primer plano, pasándole el Hwnd
            ret = SetForegroundWindow(lnHWnd)
        Else
            Set oExcel = GalenhosExcelApplication()  'New Excel.Application
            Set oWorkBook = oExcel.Workbooks.Add
            'Abre, copia y cierra la plantilla
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\ECuentaCte.xls")
            oWorkBookPlantilla.Worksheets("CuentaCte").Copy Before:=oWorkBook.Sheets(1)
            oWorkBookPlantilla.Close
            'Activa la primera hoja
            Set oWorkSheet = oWorkBook.Sheets(1)
            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
            'Inicio de Impresion
        End If
        
        If ml_IdTipoServicio = sghConsultaExterna Then
           Set rsReporte = mo_ReglasFacturacion.AtencionesFiltraIngresosPorIdAtencion(ml_idAtencion)
        Else
           Set rsReporte = mo_ReglasFacturacion.AtencionesFiltraEgresosPorIdAtencion(ml_idAtencion)
        End If
        If rsReporte.RecordCount = 0 Then
           MousePointer = 1
           Exit Sub
        End If
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(2, 0).setFormula("Estado de Cuenta al " & lcBuscaParametro.RetornaFechaServidorSQL())
            Call Feuille.getcellbyposition(2, 2).setFormula(Trim(Str(ml_idCuentaAtencion)) & "  " & txtEstadoCuenta.Text)
            Call Feuille.getcellbyposition(7, 2).setFormula(Trim(Str(ml_idAtencion)) & IIf(txtDxEgr.Text = "", "", "      Dx Egreso: " & txtDxEgr.Text))
            Call Feuille.getcellbyposition(2, 3).setFormula(txtPaciente.Text)
            Call Feuille.getcellbyposition(7, 3).setFormula(Trim(txtNroHistoria.Text) & "       Dom.Pac: " & Trim(txtDomicilioPacienteEnAtencion.Text))
            Call Feuille.getcellbyposition(2, 4).setFormula("'" & Format(rsReporte.Fields!FechaIngreso, sighentidades.DevuelveFechaSoloFormato_DMY) & " " & rsReporte.Fields!HoraIngreso)
            Call Feuille.getcellbyposition(7, 4).setFormula(IIf(IsNull(rsReporte.Fields!codServicio), "", rsReporte.Fields!codServicio & " - " & rsReporte.Fields!DServicio))
            Call Feuille.getcellbyposition(2, 5).setFormula(IIf(IsNull(rsReporte.Fields!FechaEgreso), "", "'" & Format(rsReporte.Fields!FechaEgreso & " " & rsReporte.Fields!HoraEgreso, sighentidades.DevuelveFechaSoloFormato_DMY_HM)))
            Call Feuille.getcellbyposition(7, 5).setFormula(IIf(IsNull(rsReporte.Fields!codCama), "", rsReporte.Fields!codCama))
        Else
            oWorkSheet.Cells(1, 3).Value = "Estado de Cuenta al " & lcBuscaParametro.RetornaFechaServidorSQL()
            oWorkSheet.Cells(3, 3).Value = Trim(Str(ml_idCuentaAtencion)) & "  " & txtEstadoCuenta.Text
            oWorkSheet.Cells(3, 8).Value = Trim(Str(ml_idAtencion)) & IIf(txtDxEgr.Text = "", "", "      Dx Egreso: " & txtDxEgr.Text)
            oWorkSheet.Cells(4, 3).Value = txtPaciente.Text
            oWorkSheet.Cells(4, 8).Value = Trim(txtNroHistoria.Text) & "       Dom.Pac: " & Trim(txtDomicilioPacienteEnAtencion.Text)
            oWorkSheet.Cells(5, 3).Value = "'" & Format(rsReporte.Fields!FechaIngreso, sighentidades.DevuelveFechaSoloFormato_DMY) & " " & rsReporte.Fields!HoraIngreso
            oWorkSheet.Cells(5, 8).Value = IIf(IsNull(rsReporte.Fields!codServicio), "", rsReporte.Fields!codServicio & " - " & rsReporte.Fields!DServicio)
            oWorkSheet.Cells(6, 3).Value = IIf(IsNull(rsReporte.Fields!FechaEgreso), "", "'" & Format(rsReporte.Fields!FechaEgreso & " " & rsReporte.Fields!HoraEgreso, sighentidades.DevuelveFechaSoloFormato_DMY_HM))
            oWorkSheet.Cells(6, 8).Value = IIf(IsNull(rsReporte.Fields!codCama), "", rsReporte.Fields!codCama)
        End If
        iFila = 8
        iCol = 2
        ms_EstadosFacturacion = ""
        ms_TiposFinanciamiento = ""
        ml_AgruparPor = 1
        lnTotal = 0: lnTotalSIS = 0: lnTotalSOAT = 0: lnTotalEXO = 0: lnTotalPAGO = 0: lnTotalDEBE = 0: lnTotalSALDO = 0: lnTotalCredito = 0
        
        lnTDebe = 0: lnTPago = 0: lnTSaldo = 0
        
        'Farmacia
        Set rsReporte = ucFacturacionBienes.FacturacionProductos
        If rsReporte.RecordCount > 0 Then
            rsReporte.Sort = "idServicioDeEstancia"
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF
                lnTSubTotal = 0: lnTSubTotalSIS = 0: lnTSubTotalSOAT = 0: lnTSubTotalEXO = 0: lnTsubTotalConvenio = 0
                lnTSubTotalPAGO = 0: lnTSubTotalDEBE = 0: lnTSubTotalSALDO = 0
                idPuntoCarga = rsReporte.Fields("idServicioDeEstancia").Value
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(1, iFila - 1).setFormula("(Farmacia) " & rsReporte.Fields!ServicioDeEstancia)
                Else
                    oWorkSheet.Cells(iFila, 2).Value = "(Farmacia) " & rsReporte.Fields!ServicioDeEstancia
                End If
                iFila = iFila + 1
                Do While Not rsReporte.EOF And idPuntoCarga = rsReporte.Fields("idServicioDeEstancia").Value
                   If rsReporte.Fields!idestadofacturacion = 4 Or rsReporte.Fields!idestadofacturacion = 1 Then   'Solo PAGADOS y REGISTRADOS
                        lnSIS = IIf(IsNull(rsReporte.Fields!ImporteSIS), 0, rsReporte.Fields!ImporteSIS)
                        lnEXO = IIf(IsNull(rsReporte.Fields!importeEXO), 0, rsReporte.Fields!importeEXO)
                        lnSOAT = IIf(IsNull(rsReporte.Fields!ImporteSOAT), 0, rsReporte.Fields!ImporteSOAT)
                        lnConvenio = IIf(IsNull(rsReporte.Fields!ImporteConv), 0, rsReporte.Fields!ImporteConv)
                        
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(2, iFila - 1).setFormula(rsReporte.Fields("Codigo").Value & " - " & rsReporte.Fields("NombreProducto").Value)
                            Call Feuille.getcellbyposition(4, iFila - 1).setFormula(rsReporte.Fields("CantidadPagar").Value)
                            Call Feuille.getcellbyposition(5, iFila - 1).setFormula(rsReporte.Fields("preciounitario").Value)
                            Call Feuille.getcellbyposition(6, iFila - 1).setFormula(Format(rsReporte.Fields("TotalPagar").Value, "####,###.#0"))
                            Call Feuille.getcellbyposition(7, iFila - 1).setFormula(Format(lnEXO, "####,###.#0"))
                            Call Feuille.getcellbyposition(8, iFila - 1).setFormula(Format(lnSIS, "####,###.#0"))
                            Call Feuille.getcellbyposition(9, iFila - 1).setFormula(Format(lnSOAT, "####,###.#0"))
                            Call Feuille.getcellbyposition(10, iFila - 1).setFormula(Format(lnConvenio, "####,###.#0"))
                        Else
                            oWorkSheet.Cells(iFila, 3).Value = rsReporte.Fields("Codigo").Value & " - " & rsReporte.Fields("NombreProducto").Value
                            oWorkSheet.Cells(iFila, 5).Value = rsReporte.Fields("CantidadPagar").Value
                            oWorkSheet.Cells(iFila, 6).Value = rsReporte.Fields("preciounitario").Value
                            oWorkSheet.Cells(iFila, 7).Value = Format(rsReporte.Fields("TotalPagar").Value, "####,###.#0")
                            oWorkSheet.Cells(iFila, 8).Value = Format(lnEXO, "####,###.#0")
                            oWorkSheet.Cells(iFila, 9).Value = Format(lnSIS, "####,###.#0")
                            oWorkSheet.Cells(iFila, 10).Value = Format(lnSOAT, "####,###.#0")
                            oWorkSheet.Cells(iFila, 11).Value = Format(lnConvenio, "####,###.#0")
                        End If

                        If lbGeneraReciboPago = True Then
                           lnDebe = rsReporte.Fields!TotalPagar - lnEXO - lnSIS - lnSOAT
                        Else
                           If (rsReporte.Fields!CantidadSIS + rsReporte.Fields!CantidadSOAT + rsReporte.Fields!cantidadConv) = 0 Then
                               lnDebe = rsReporte.Fields!TotalPagar
                           Else
                               lnDebe = 0
                           End If
                        End If
                        If rsReporte.Fields!idestadofacturacion = 4 Then
                           lnPago = rsReporte.Fields("ImporteEnBoleta").Value - lnEXO
                        Else
                           lnPago = 0
                        End If
                        lnSaldo = lnDebe - lnPago
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnDebe, "####,###.#0"))
                            Call Feuille.getcellbyposition(12, iFila - 1).setFormula(Format(lnPago, "####,###.#0"))
                            Call Feuille.getcellbyposition(13, iFila - 1).setFormula(Format(lnSaldo, "####,###.#0"))
                            Call Feuille.getcellbyposition(14, iFila - 1).setFormula(rsReporte.Fields!nroDcto)
                            Call Feuille.getcellbyposition(15, iFila - 1).setFormula(rsReporte.Fields!FechaDespacho)
                        Else
                            oWorkSheet.Cells(iFila, 12).Value = Format(lnDebe, "####,###.#0")
                            oWorkSheet.Cells(iFila, 13).Value = Format(lnPago, "####,###.#0")
                            oWorkSheet.Cells(iFila, 14).Value = Format(lnSaldo, "####,###.#0")
                            oWorkSheet.Cells(iFila, 15).Value = rsReporte.Fields!nroDcto
                            oWorkSheet.Cells(iFila, 16).Value = rsReporte.Fields!FechaDespacho
                        End If
                        lnTSubTotal = lnTSubTotal + rsReporte.Fields("TotalPagar").Value
                        lnTSubTotalEXO = lnTSubTotalEXO + lnEXO
                        lnTSubTotalSIS = lnTSubTotalSIS + lnSIS
                        lnTSubTotalSOAT = lnTSubTotalSOAT + lnSOAT
                        lnTsubTotalConvenio = lnTsubTotalConvenio + lnConvenio
                        lnTSubTotalPAGO = lnTSubTotalPAGO + lnPago
                        lnTSubTotalDEBE = lnTSubTotalDEBE + lnDebe
                        lnTSubTotalSALDO = lnTSubTotalSALDO + lnSaldo
                        
                        lnTotal = lnTotal + rsReporte.Fields("TotalPagar").Value
                        lnTotalSIS = lnTotalSIS + lnSIS
                        lnTotalSOAT = lnTotalSOAT + lnSOAT
                        lnTotalEXO = lnTotalEXO + lnEXO
                        lnTotalConvenio = lnTotalConvenio + lnConvenio
                        lnTotalPAGO = lnTotalPAGO + lnPago
                        lnTotalDEBE = lnTotalDEBE + lnDebe
                        lnTotalSALDO = lnTotalSALDO + lnSaldo
                
                        If rsReporte.Fields!idProducto = lnIdPagosACuenta Then   'Pagos a cuenta
                           lnTotalCredito = lnTotalCredito + rsReporte.Fields("TotalPorPagar").Value
                        End If
                        
                        iFila = iFila + 1
                    End If
                    rsReporte.MoveNext
                    If rsReporte.EOF Then
                       Exit Do
                    End If
                 Loop
                 If lbEsOpenOffice = True Then
                    Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(3) & CStr(iFila) & ":" & mo_ReglasReportes.BuscaNombreColumna(iCol + 14) & CStr(iFila))
                    mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                 Else
                     mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 3, iFila, iCol + 14
                 End If
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(6, iFila - 1).setFormula(Format(lnTSubTotal, "####,###.#0"))
                    Call Feuille.getcellbyposition(7, iFila - 1).setFormula(Format(lnTSubTotalEXO, "####,###.#0"))
                    Call Feuille.getcellbyposition(8, iFila - 1).setFormula(Format(lnTSubTotalSIS, "####,###.#0"))
                    Call Feuille.getcellbyposition(9, iFila - 1).setFormula(Format(lnTSubTotalSOAT, "####,###.#0"))
                    Call Feuille.getcellbyposition(10, iFila - 1).setFormula(Format(lnTsubTotalConvenio, "####,###.#0"))
                    Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnTSubTotalDEBE, "####,###.#0"))
                    Call Feuille.getcellbyposition(12, iFila - 1).setFormula(Format(lnTSubTotalPAGO, "####,###.#0"))
                    Call Feuille.getcellbyposition(13, iFila - 1).setFormula(Format(lnTSubTotalSALDO, "####,###.#0"))
                Else
                     oWorkSheet.Cells(iFila, 7).Value = Format(lnTSubTotal, "####,###.#0")
                     oWorkSheet.Cells(iFila, 8).Value = Format(lnTSubTotalEXO, "####,###.#0")
                     oWorkSheet.Cells(iFila, 9).Value = Format(lnTSubTotalSIS, "####,###.#0")
                     oWorkSheet.Cells(iFila, 10).Value = Format(lnTSubTotalSOAT, "####,###.#0")
                     oWorkSheet.Cells(iFila, 11).Value = Format(lnTsubTotalConvenio, "####,###.#0")
                     oWorkSheet.Cells(iFila, 12).Value = Format(lnTSubTotalDEBE, "####,###.#0")
                     oWorkSheet.Cells(iFila, 13).Value = Format(lnTSubTotalPAGO, "####,###.#0")
                     oWorkSheet.Cells(iFila, 14).Value = Format(lnTSubTotalSALDO, "####,###.#0")
                End If
                 iFila = iFila + 1
             Loop
        End If
        rsReporte.Close
        lnPagoEnFarmacia = lnTSubTotalSALDO
        'Servicios
        lnTotalPagarEstancia = 0
        lnFor = 1
        If ml_IdTipoServicio = sghHospitalizacion Then
           lnFor = 2
        End If
        For lnForNum = 1 To lnFor
            If lnForNum = 1 Then
                Set rsReporte = ucFacturacionServicios.FacturacionProductos
                If ml_IdTipoServicio = sghHospitalizacion Then
                   On Error Resume Next
                   rsReporte.Filter = "idPuntoCarga<>" & sghPtoCargaAdmisionHospitalizacion
                End If
            Else
                rsReporte.Filter = ""
                If txtFegreso.Text = "" Then
                    ldFechaAlta = CDate(Format(Now, sighentidades.DevuelveFechaSoloFormato_DMY))
                    lcHoraAlta = Format(Now, sighentidades.DevuelveHoraSoloFormato_HM)
                Else
                    ldFechaAlta = CDate(Format(CDate(txtFegreso.Text), sighentidades.DevuelveFechaSoloFormato_DMY))
                    lcHoraAlta = Format(CDate(txtFegreso.Text), sighentidades.DevuelveHoraSoloFormato_HM)
                End If
                oGenerarRecordsetProductos.GenerarRecordsetProductos rsReporte
                mo_AdminAdmision.GeneraEstanciaPorCadaServicioHospitalizado ml_idCuentaAtencion, ldFechaAlta, lcHoraAlta, rsReporte, lnTotalPagarEstancia, lnTotalDiasEstancia, oConexion
                If txtFegreso.Text <> "" Then
                   lnTotalPagarEstancia = 0
                End If
            End If
            'Set rsReporte = ucFacturacionServicios.FacturacionProductos
            lnTotalPagosAdelantados = 0
            If rsReporte.RecordCount > 0 Then
                rsReporte.Sort = "idServicioDeEstancia"
                rsReporte.MoveFirst
                Do While Not rsReporte.EOF
                    idPuntoCarga = rsReporte.Fields("idServicioDeEstancia").Value
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(1, iFila - 1).setFormula("(Servicios) " & rsReporte.Fields!ServicioDeEstancia)
                    Else
                        oWorkSheet.Cells(iFila, 2).Value = "(Servicios) " & rsReporte.Fields!ServicioDeEstancia
                    End If
                    lnTSubTotal = 0: lnTSubTotalSIS = 0: lnTSubTotalSOAT = 0: lnTSubTotalEXO = 0: lnTsubTotalConvenio = 0
                    lnTSubTotalPAGO = 0: lnTSubTotalDEBE = 0: lnTSubTotalSALDO = 0
                    iFila = iFila + 1
                    Do While Not rsReporte.EOF And idPuntoCarga = rsReporte.Fields("idServicioDeEstancia").Value
                        If rsReporte.Fields!idestadofacturacion = 4 Or rsReporte.Fields!idestadofacturacion = 1 Or rsReporte.Fields!idestadofacturacion = 10 Or rsReporte.Fields!idestadofacturacion = sghConPreVenta Then  'Solo PAGADOS/REGISTRADOS/AUTORIZ.AUTOMATICA/preventa
                            lnSIS = IIf(IsNull(rsReporte.Fields!ImporteSIS), 0, rsReporte.Fields!ImporteSIS)
                            lnSOAT = IIf(IsNull(rsReporte.Fields!ImporteSOAT), 0, rsReporte.Fields!ImporteSOAT)
                            lnEXO = IIf(IsNull(rsReporte.Fields!importeEXO), 0, rsReporte.Fields!importeEXO)
                            lnConvenio = IIf(IsNull(rsReporte.Fields!ImporteConv), 0, rsReporte.Fields!ImporteConv)
                            
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(2, iFila - 1).setFormula(rsReporte.Fields("Codigo").Value & " - " & rsReporte.Fields("NombreProducto").Value)
                                Call Feuille.getcellbyposition(4, iFila - 1).setFormula(rsReporte.Fields("CantidadPagar").Value)
                                Call Feuille.getcellbyposition(5, iFila - 1).setFormula(Format(rsReporte.Fields("preciounitario").Value, "####,###.###0"))
                                Call Feuille.getcellbyposition(6, iFila - 1).setFormula(Format(rsReporte.Fields("TotalPagar").Value, "####,###.#0"))
                                Call Feuille.getcellbyposition(7, iFila - 1).setFormula(Format(lnEXO, "####,###.#0"))
                                Call Feuille.getcellbyposition(8, iFila - 1).setFormula(Format(lnSIS, "####,###.#0"))
                                Call Feuille.getcellbyposition(9, iFila - 1).setFormula(Format(lnSOAT, "####,###.#0"))
                                Call Feuille.getcellbyposition(10, iFila - 1).setFormula(Format(lnConvenio, "####,###.#0"))
                            Else
                                oWorkSheet.Cells(iFila, 3).Value = rsReporte.Fields("Codigo").Value & " - " & rsReporte.Fields("NombreProducto").Value
                                oWorkSheet.Cells(iFila, 5).Value = rsReporte.Fields("CantidadPagar").Value
                                oWorkSheet.Cells(iFila, 6).Value = rsReporte.Fields("preciounitario").Value
                                oWorkSheet.Cells(iFila, 7).Value = Format(rsReporte.Fields("TotalPagar").Value, "####,###.#0")
                                oWorkSheet.Cells(iFila, 8).Value = Format(lnEXO, "####,###.#0")
                                oWorkSheet.Cells(iFila, 9).Value = Format(lnSIS, "####,###.#0")
                                oWorkSheet.Cells(iFila, 10).Value = Format(lnSOAT, "####,###.#0")
                                oWorkSheet.Cells(iFila, 11).Value = Format(lnConvenio, "####,###.#0")
                            End If
                            
                            If rsReporte.Fields!idestadofacturacion = 4 Then
                                lnPago = rsReporte.Fields("ImporteEnBoleta").Value - lnEXO
                            Else
                                lnPago = 0
                            End If
                            If rsReporte.Fields!idProducto <> lnIdPagosACuenta And rsReporte.Fields!idProducto <> lnIdPagosXdevoluciones Then
                               If lbGeneraReciboPago = True Then
                                    lnDebe = rsReporte.Fields!TotalPagar - lnEXO - lnSIS - lnSOAT
                               Else
                                    If (rsReporte.Fields!CantidadSIS + rsReporte.Fields!CantidadSOAT + rsReporte.Fields!cantidadConv) = 0 Then
                                        lnDebe = rsReporte.Fields!TotalPagar
                                    Else
                                        lnDebe = 0
                                    End If
                               End If
                               lnSaldo = lnDebe - lnPago
                            Else
                               If rsReporte.Fields!idProducto = lnIdPagosACuenta Then
                                    lnTotalPagosAdelantados = lnTotalPagosAdelantados + rsReporte.Fields!ImporteEnBoleta
                                    lnDebe = 0
                                    If lbGeneraReciboPago = True Then
                                       lnSaldo = -rsReporte.Fields!ImporteEnBoleta
                                    Else
                                       lnSaldo = 0
                                    End If
                               Else
                                    lnDebe = 0
                                    If lbGeneraReciboPago = True Then
                                       lnSaldo = rsReporte.Fields!ImporteEnBoleta
                                       lnPago = -rsReporte.Fields!ImporteEnBoleta
                                    Else
                                       lnSaldo = 0
                                       lnPago = -rsReporte.Fields!ImporteEnBoleta
                                    End If
                               End If
                            End If
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnDebe, "####,###.#0"))
                                Call Feuille.getcellbyposition(12, iFila - 1).setFormula(Format(lnPago, "####,###.#0"))
                                Call Feuille.getcellbyposition(13, iFila - 1).setFormula(Format(lnSaldo, "####,###.#0"))
                                Call Feuille.getcellbyposition(14, iFila - 1).setFormula(rsReporte.Fields!nroDcto)
                                Call Feuille.getcellbyposition(15, iFila - 1).setFormula(rsReporte.Fields!FechaDespacho)
                            Else
                                oWorkSheet.Cells(iFila, 12).Value = Format(lnDebe, "####,###.#0")
                                oWorkSheet.Cells(iFila, 13).Value = Format(lnPago, "####,###.#0")
                                oWorkSheet.Cells(iFila, 14).Value = Format(lnSaldo, "####,###.#0")
                                oWorkSheet.Cells(iFila, 15).Value = rsReporte.Fields!nroDcto
                                oWorkSheet.Cells(iFila, 16).Value = rsReporte.Fields!FechaDespacho
                            End If
    
                            If rsReporte.Fields!idProducto <> lnIdPagosACuenta Then
                               lnTSubTotal = lnTSubTotal + rsReporte.Fields("TotalPagar").Value
                               lnTSubTotalDEBE = lnTSubTotalDEBE + lnDebe
                            End If
                            lnTSubTotalEXO = lnTSubTotalEXO + lnEXO
                            lnTSubTotalSIS = lnTSubTotalSIS + lnSIS
                            lnTSubTotalSOAT = lnTSubTotalSOAT + lnSOAT
                            lnTsubTotalConvenio = lnTsubTotalConvenio + lnConvenio
                            lnTSubTotalPAGO = lnTSubTotalPAGO + lnPago
                            lnTSubTotalSALDO = lnTSubTotalSALDO + lnSaldo
                            
                            If rsReporte.Fields!idProducto <> lnIdPagosACuenta Then
                               lnTotal = lnTotal + rsReporte.Fields("TotalPagar").Value
                               lnTotalDEBE = lnTotalDEBE + lnDebe
                            End If
                            lnTotalSIS = lnTotalSIS + lnSIS
                            lnTotalSOAT = lnTotalSOAT + lnSOAT
                            lnTotalEXO = lnTotalEXO + lnEXO
                            lnTotalConvenio = lnTotalConvenio + lnConvenio
                            lnTotalPAGO = lnTotalPAGO + lnPago
                            lnTotalSALDO = lnTotalSALDO + lnSaldo
                            
                            If rsReporte.Fields!idProducto = lnIdPagosACuenta Then   'Pagos a cuenta
                               lnTotalCredito = lnTotalCredito + rsReporte.Fields("TotalPorPagar").Value
                            End If
                            
                            iFila = iFila + 1
                        End If
                        rsReporte.MoveNext
                        If rsReporte.EOF Then
                           Exit Do
                        End If
                    Loop
                    If lbEsOpenOffice = True Then
                        Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(3) & CStr(iFila) & ":" & mo_ReglasReportes.BuscaNombreColumna(iCol + 14) & CStr(iFila))
                        mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                    Else
                        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 3, iFila, iCol + 14
                    End If
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(6, iFila - 1).setFormula(Format(lnTSubTotal, "####,###.#0"))
                        Call Feuille.getcellbyposition(7, iFila - 1).setFormula(Format(lnTSubTotalEXO, "####,###.#0"))
                        Call Feuille.getcellbyposition(8, iFila - 1).setFormula(Format(lnTSubTotalSIS, "####,###.#0"))
                        Call Feuille.getcellbyposition(9, iFila - 1).setFormula(Format(lnTSubTotalSOAT, "####,###.#0"))
                        Call Feuille.getcellbyposition(10, iFila - 1).setFormula(Format(lnTsubTotalConvenio, "####,###.#0"))
                        Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnTSubTotalDEBE, "####,###.#0"))
                        Call Feuille.getcellbyposition(12, iFila - 1).setFormula(Format(lnTSubTotalPAGO, "####,###.#0"))
                        Call Feuille.getcellbyposition(13, iFila - 1).setFormula(Format(lnTSubTotalSALDO, "####,###.#0"))
                    Else
                        oWorkSheet.Cells(iFila, 7).Value = Format(lnTSubTotal, "####,###.#0")
                        oWorkSheet.Cells(iFila, 8).Value = Format(lnTSubTotalEXO, "####,###.#0")
                        oWorkSheet.Cells(iFila, 9).Value = Format(lnTSubTotalSIS, "####,###.#0")
                        oWorkSheet.Cells(iFila, 10).Value = Format(lnTSubTotalSOAT, "####,###.#0")
                        oWorkSheet.Cells(iFila, 11).Value = Format(lnTsubTotalConvenio, "####,###.#0")
                        oWorkSheet.Cells(iFila, 12).Value = Format(lnTSubTotalDEBE, "####,###.#0")
                        oWorkSheet.Cells(iFila, 13).Value = Format(lnTSubTotalPAGO, "####,###.#0")
                        oWorkSheet.Cells(iFila, 14).Value = Format(lnTSubTotalSALDO, "####,###.#0")
                    End If
                    iFila = iFila + 1
                 Loop
            End If
            iFila = iFila + 1
        Next
        If lbEsOpenOffice = True Then
            Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(2) & CStr(iFila) & ":" & mo_ReglasReportes.BuscaNombreColumna(iCol + 14) & CStr(iFila))
            mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
        Else
            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, iCol + 14
        End If
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Total: ")
            Call Feuille.getcellbyposition(6, iFila - 1).setFormula(Format(lnTotal, "####,###.#0"))
            Call Feuille.getcellbyposition(7, iFila - 1).setFormula(Format(lnTotalEXO, "####,###.#0"))
            Call Feuille.getcellbyposition(8, iFila - 1).setFormula(Format(lnTotalSIS, "####,###.#0"))
            Call Feuille.getcellbyposition(9, iFila - 1).setFormula(Format(lnTotalSOAT, "####,###.#0"))
            Call Feuille.getcellbyposition(10, iFila - 1).setFormula(Format(lnTotalConvenio, "####,###.#0"))
            Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnTotalDEBE, "####,###.#0"))
            Call Feuille.getcellbyposition(12, iFila - 1).setFormula(Format(lnTotalPAGO, "####,###.#0"))
            Call Feuille.getcellbyposition(13, iFila - 1).setFormula(Format(lnTotalSALDO, "####,###.#0"))
        Else
            oWorkSheet.Cells(iFila, 2).Value = "Total: "
            oWorkSheet.Cells(iFila, 7).Value = Format(lnTotal, "####,###.#0")
            oWorkSheet.Cells(iFila, 8).Value = Format(lnTotalEXO, "####,###.#0")
            oWorkSheet.Cells(iFila, 9).Value = Format(lnTotalSIS, "####,###.#0")
            oWorkSheet.Cells(iFila, 10).Value = Format(lnTotalSOAT, "####,###.#0")
            oWorkSheet.Cells(iFila, 11).Value = Format(lnTotalConvenio, "####,###.#0")
            oWorkSheet.Cells(iFila, 12).Value = Format(lnTotalDEBE, "####,###.#0")
            oWorkSheet.Cells(iFila, 13).Value = Format(lnTotalPAGO, "####,###.#0")
            oWorkSheet.Cells(iFila, 14).Value = Format(lnTotalSALDO, "####,###.#0")
        End If
        If lbGeneraReciboPago = True Then
            If lnTotalPagosAdelantados > lnPagoEnFarmacia Then
               lnTotalPagosAdelantados = lnTotalPagosAdelantados - lnPagoEnFarmacia
               lnPagoEnFarmacia = 0
            Else
               lnPagoEnFarmacia = lnPagoEnFarmacia - lnTotalPagosAdelantados
               lnTotalPagosAdelantados = 0
            End If
            lnPagoEnServicios = lnPagoEnServicios - lnTotalPagosAdelantados
        Else
           Select Case mo_ReglasFacturacion.TiposFinanciamientoDevuelveComoSeTrabajaEnEstadoCuenta(lnIdTipoFinanciamientoActual, oConexion)
           Case sghTrabajaSeguroSIS
               lnTotalSIS = lnTotalSIS - lnTotalPagosAdelantados
           Case sghTrabajaSeguroSOAT
               lnTotalSOAT = lnTotalSOAT - lnTotalPagosAdelantados
           Case sghTrabajaSeguroConvenios
               lnTotalConvenio = lnTotalConvenio - lnTotalPagosAdelantados
           End Select
        End If
        lnPagoEnServicios = lnTotalSALDO - lnPagoEnFarmacia
        
        iFila = iFila + 3
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(10, iFila - 1).setFormula("TOTAL CUENTA")
            Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnTotal, "####,###.#0"))
            iFila = iFila + 1
            Call Feuille.getcellbyposition(10, iFila - 1).setFormula("EXONERADO")
            Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnTotalEXO, "####,###.#0"))
            iFila = iFila + 1
            Call Feuille.getcellbyposition(10, iFila - 1).setFormula("SIS CUBRE (-PAGOS A CUENTA)")
            Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnTotalSIS, "####,###.#0"))
            iFila = iFila + 1
            Call Feuille.getcellbyposition(10, iFila - 1).setFormula("SOAT CUBRE (-PAGOS A CUENTA)")
            Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnTotalSOAT, "####,###.#0"))
            iFila = iFila + 1
            Call Feuille.getcellbyposition(10, iFila - 1).setFormula("CONVENIOS CUBRE (-PAGOS A CUENTA)")
            Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnTotalConvenio, "####,###.#0"))
            iFila = iFila + 1
            Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(10) & CStr(iFila) & ":" & mo_ReglasReportes.BuscaNombreColumna(iCol + 10) & CStr(iFila))
            mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
            Call Feuille.getcellbyposition(10, iFila - 1).setFormula("TOTAL DEUDA")
            Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnTotalDEBE, "####,###.#0"))
            iFila = iFila + 1
            Call Feuille.getcellbyposition(10, iFila - 1).setFormula("PAGOS REALIZADOS")
            Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnTotalPAGO, "####,###.#0"))
            iFila = iFila + 1
            Call Feuille.getcellbyposition(10, iFila - 1).setFormula("CAJA DEBE INGRESAR")
            iFila = iFila + 1
            Call Feuille.getcellbyposition(10, iFila - 1).setFormula("CREDITO")
            Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnTotalCredito, "####,###.#0"))
            iFila = iFila + 1
            Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(10) & CStr(iFila) & ":" & mo_ReglasReportes.BuscaNombreColumna(iCol + 10) & CStr(iFila))
            mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
        Else
            oWorkSheet.Cells(iFila, 11).Value = "TOTAL CUENTA"
            oWorkSheet.Cells(iFila, 12).Value = Format(lnTotal, "####,###.#0")
            iFila = iFila + 1
            oWorkSheet.Cells(iFila, 11).Value = "EXONERADO"
            oWorkSheet.Cells(iFila, 12).Value = Format(lnTotalEXO, "####,###.#0")
            iFila = iFila + 1
            oWorkSheet.Cells(iFila, 11).Value = "SIS CUBRE (-PAGOS A CUENTA)"
            oWorkSheet.Cells(iFila, 12).Value = Format(lnTotalSIS, "####,###.#0")
            iFila = iFila + 1
            oWorkSheet.Cells(iFila, 11).Value = "SOAT CUBRE (-PAGOS A CUENTA)"
            oWorkSheet.Cells(iFila, 12).Value = Format(lnTotalSOAT, "####,###.#0")
            iFila = iFila + 1
            oWorkSheet.Cells(iFila, 11).Value = "CONVENIOS CUBRE (-PAGOS A CUENTA)"
            oWorkSheet.Cells(iFila, 12).Value = Format(lnTotalConvenio, "####,###.#0")
            iFila = iFila + 1
            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 10, iFila, iCol + 10
            oWorkSheet.Cells(iFila, 11).Value = "TOTAL DEUDA"
            oWorkSheet.Cells(iFila, 12).Value = Format(lnTotalDEBE, "####,###.#0")
            iFila = iFila + 1
            oWorkSheet.Cells(iFila, 11).Value = "PAGOS REALIZADOS"
            oWorkSheet.Cells(iFila, 12).Value = Format(lnTotalPAGO, "####,###.#0")
            iFila = iFila + 1
            oWorkSheet.Cells(iFila, 11).Value = "CAJA DEBE INGRESAR"
            iFila = iFila + 1
            oWorkSheet.Cells(iFila, 11).Value = "CREDITO"
            oWorkSheet.Cells(iFila, 12).Value = Format(lnTotalCredito, "####,###.#0")
            iFila = iFila + 1
            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 10, iFila, iCol + 10
        End If
        
        lnPagoEnServicios = CCur(txtTotalServicios.Text) + lnTotalPagarEstancia
        lnPagoEnFarmacia = CCur(txtTotalFarmacia.Text)
        lnTotalSALDO = lnPagoEnServicios + lnPagoEnFarmacia
        
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(10, iFila - 1).setFormula("PACIENTE DEBE PAGAR")
            Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnTotalSALDO, "####,###.#0"))
            iFila = iFila + 1
            Call Feuille.getcellbyposition(10, iFila - 1).setFormula("PAGO POR CONSUMO FARMACIA")
            Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnPagoEnFarmacia, "####,###.#0"))
            iFila = iFila + 1
            Call Feuille.getcellbyposition(10, iFila - 1).setFormula("PAGO POR CONSUMO SERVICIO")
            Call Feuille.getcellbyposition(11, iFila - 1).setFormula(Format(lnPagoEnServicios, "####,###.#0"))
        Else
            oWorkSheet.Cells(iFila, 11).Value = "PACIENTE DEBE PAGAR"
            oWorkSheet.Cells(iFila, 12).Value = Format(lnTotalSALDO, "####,###.#0")
            iFila = iFila + 1
            oWorkSheet.Cells(iFila, 11).Value = "PAGO POR CONSUMO FARMACIA"
            oWorkSheet.Cells(iFila, 12).Value = Format(lnPagoEnFarmacia, "####,###.#0")
            iFila = iFila + 1
            oWorkSheet.Cells(iFila, 11).Value = "PAGO POR CONSUMO SERVICIO"
            oWorkSheet.Cells(iFila, 12).Value = Format(lnPagoEnServicios, "####,###.#0")
        End If
        
        'Transferencias
        rsReporte.Close
        Set rsReporte = mo_ReglasFacturacion.AtencionesFiltraTransferenciasPorIdAtencion(ml_idAtencion)
        If rsReporte.RecordCount > 0 Then
            iFila = iFila - rsReporte.RecordCount - 1
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("ESTADIA")
            Else
                oWorkSheet.Cells(iFila, 2).Value = "ESTADIA"
            End If
            iFila = iFila + 1
            If lbEsOpenOffice = True Then
                Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(2) & CStr(iFila) & ":" & mo_ReglasReportes.BuscaNombreColumna(6) & CStr(iFila))
                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Cod.Cama")
                Call Feuille.getcellbyposition(2, iFila - 1).setFormula("Servicio que fue transferido")
                Call Feuille.getcellbyposition(4, iFila - 1).setFormula("F.Transf")
                Call Feuille.getcellbyposition(5, iFila - 1).setFormula("H.Transf")
            Else
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 6
                oWorkSheet.Cells(iFila, 2).Value = "Cod.Cama"
                oWorkSheet.Cells(iFila, 3).Value = "Servicio que fue transferido"
                oWorkSheet.Cells(iFila, 5).Value = "F.Transf"
                oWorkSheet.Cells(iFila, 6).Value = "H.Transf"
            End If
            iFila = iFila + 1
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF
                If lbEsOpenOffice = True Then
                    'Call Feuille.getcellbyposition(1, iFila - 1).setFormula(rsReporte!CodigoCama)
                    'Yamill palomino 15/10/2014
                    Call Feuille.getcellbyposition(1, iFila - 1).setFormula(IIf(IsNull(rsReporte!CodigoCama), "", (rsReporte!CodigoCama)))
                    Call Feuille.getcellbyposition(2, iFila - 1).setFormula(rsReporte!NombreServicio)
                    Call Feuille.getcellbyposition(4, iFila - 1).setFormula("'" & rsReporte!FechaOcupacion)
                    Call Feuille.getcellbyposition(5, iFila - 1).setFormula(rsReporte!HoraOcupacion)
                Else
                    oWorkSheet.Cells(iFila, 2).Value = rsReporte!CodigoCama
                    oWorkSheet.Cells(iFila, 3).Value = rsReporte!NombreServicio
                    oWorkSheet.Cells(iFila, 5).Value = "'" & rsReporte!FechaOcupacion
                    oWorkSheet.Cells(iFila, 6).Value = rsReporte!HoraOcupacion
                    'oWorkSheet.Cells(iFila, 7).Value = rsReporte!NombreMedico
                End If
                iFila = iFila + 1
                rsReporte.MoveNext
            Loop
        End If
        '***Donaciones en Farmacia
        If rsItemsDonaciones.RecordCount > 0 Then
           Dim lnCantidadDona As Long, lcCodigoDona As String, lbContinua As Boolean
           iFila = iFila + 2
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("LISTA DE DONACIONES:")
            Else
                oWorkSheet.Cells(iFila, 2).Value = "LISTA DE DONACIONES:"
            End If
            iFila = iFila + 1
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Descripción")
                Call Feuille.getcellbyposition(8, iFila - 1).setFormula("Cantidad")
            Else
                oWorkSheet.Cells(iFila, 2).Value = "Descripción"
                oWorkSheet.Cells(iFila, 9).Value = "Cantidad"
            End If
            If lbEsOpenOffice = True Then
                Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(2) & CStr(iFila) & ":" & mo_ReglasReportes.BuscaNombreColumna(9) & CStr(iFila))
                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
            Else
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 9
            End If
           Set rsReporte = Nothing
           With rsReporte
                  .Fields.Append "Codigo", adVarChar, 10, adFldIsNullable
                  .Fields.Append "Nombre", adVarChar, 200, adFldIsNullable
                  .Fields.Append "Cantidad", adInteger
                  .CursorType = adOpenDynamic
                  .LockType = adLockOptimistic
                  .Open
           End With
           rsItemsDonaciones.MoveFirst
           Do While Not rsItemsDonaciones.EOF
              lbContinua = True
              If rsReporte.RecordCount > 0 Then
                 rsReporte.MoveFirst
                 rsReporte.Find "Codigo='" & rsItemsDonaciones.Fields!Codigo & "'"
                 If Not rsReporte.EOF Then
                    lbContinua = False
                 End If
              End If
              If lbContinua = True Then
                  rsReporte.AddNew
                  rsReporte.Fields!Codigo = rsItemsDonaciones.Fields!Codigo
                  rsReporte.Fields!nombre = rsItemsDonaciones.Fields!nombre
              End If
              rsReporte.Fields!cantidad = rsReporte.Fields!cantidad + rsItemsDonaciones.Fields!cantidad
              rsReporte.Update
              rsItemsDonaciones.MoveNext
           Loop
           rsReporte.Sort = "nombre,codigo"
           rsReporte.MoveFirst
           Do While Not rsReporte.EOF
              iFila = iFila + 1
              lcCodigoDona = rsReporte.Fields!Codigo
              If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula(Trim(rsReporte.Fields!Codigo) & " " & rsReporte.Fields!nombre)
                Call Feuille.getcellbyposition(8, iFila - 1).setFormula(rsReporte.Fields!cantidad)
              Else
                oWorkSheet.Cells(iFila, 2).Value = Trim(rsReporte.Fields!Codigo) & " " & rsReporte.Fields!nombre
                oWorkSheet.Cells(iFila, 9).Value = rsReporte.Fields!cantidad
              End If
              rsReporte.MoveNext
           Loop
        End If
        '
        If lcListaDeOrdenesDePago <> "" Then
            iFila = iFila + 2
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("* El CAJERO debe emitir Boletas usando " & lcListaDeOrdenesDePago)
                
            Else
                oWorkSheet.Cells(iFila, 2).Value = "* El CAJERO debe emitir Boletas usando " & lcListaDeOrdenesDePago
            End If
        End If
        '
        'yamill palomino 15/10/2014
        If lbEsOpenOffice = True Then
            Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
            PrintArea(0).Sheet = 0
            PrintArea(0).startcolumn = 1
            PrintArea(0).StartRow = 0
            PrintArea(0).EndColumn = 17
            PrintArea(0).EndRow = iFila
            Call Feuille.SetPrintAreas(PrintArea())
            Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
            Call Feuille.SetPrintAreas(PrintArea())
            Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
            MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
        Else
            oWorkSheet.PageSetup.PrintTitleRows = "$1:$7"
            If oWorkSheet.PageSetup.PrintArea <> "" Then oWorkSheet.PageSetup.PrintArea = "$A$1:$R$" & (iFila + 2) 'sighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
            oExcel.Visible = True
            oWorkSheet.PrintPreview
        End If
        'oWorkSheet.PrintOut
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
        'encabezado de pagina
        Set PageStyles = Nothing
        Set Sheet = Nothing
        Set StyleFamilies = Nothing
        Set DefPage = Nothing
        Set Htext = Nothing
        Set Hcontent = Nothing
    Else
        'liberar memoria
        Set oExcel = Nothing
        Set oWorkBookPlantilla = Nothing
        Set oWorkBook = Nothing
        Set oWorkSheet = Nothing
    End If
    oConexion.Close
    Set oConexion = Nothing
    MousePointer = 1

End Sub

Private Sub cmdLiquidacion_Click()
'    Dim oExcel As Excel.Application
'    Dim oWorkBookPlantilla As Workbook
'    Dim oWorkBook As Workbook
'    Dim oWorkSheet As Worksheet
'    Dim iFila As Long, iCol As Integer
'    Dim rsReporte As New Recordset
'    Dim ms_EstadosFacturacion As String
'    Dim ms_TiposFinanciamiento As String
'    Dim ml_AgruparPor As Long
'    Dim mo_ReporteUtil As New ReporteUtil
'    Dim idPuntoCarga As Long
'
'    Dim lnTSubTotal As Double: Dim lnTSubTotalSIS As Double: Dim lnTSubTotalSOAT As Double: Dim lnTSubTotalEXO As Double: Dim lnTsubTotalConvenio As Double
'    Dim lnTSubTotalPAGO As Double: Dim lnTSubTotalDEBE As Double: Dim lnTSubTotalSALDO As Double
'
'    Dim lnTotal As Double: Dim lnTotalSIS As Double: Dim lnTotalSOAT As Double: Dim lnTotalEXO As Double: Dim lnTotalConvenio As Double
'    Dim lnTotalPAGO As Double: Dim lnTotalDEBE As Double: Dim lnTotalSALDO As Double
'
'    Dim lnDebe As Double: Dim lnPago As Double: Dim lnSaldo As Double
'    Dim lnTDebe As Double: Dim lnTPago As Double: Dim lnTSaldo As Double
'    Dim lnSIS As Double: Dim lnSOAT As Double: Dim lnEXO As Double: Dim lnTotalCredito As Double: Dim lnConvenio As Double
'    Dim lnDctos As Double: Dim lnPagoEnFarmacia As Double: Dim lnPagoEnServicios As Double
'    Dim lnTotalPagosAdelantados As Double
'    Dim lnCantidadPagarBienes As Long, lnTotalPagarBienes As Double
'    Dim lcBuscaParametro As New SIGHDatos.Parametros
'    Dim lnTotalPagarEstancia As Double, lnTotalDiasEstancia As Long
'    Dim ldFechaAlta As Date, lcHoraAlta As String, lcCabeceraLiquidacion As String
'    Dim oConexion As New Connection
'    oConexion.Open SIGHEntidades.CadenaConexion
'    oConexion.CursorLocation = adUseClient
'    If txtPaciente.Text = "" Then
'        MsgBox "Tiene que LEER el Estado de Cuenta del Paciente", vbInformation, "Resultado"
'    Else
'        If ucFacturacionBienes.FacturacionProductos.RecordCount = 0 And ucFacturacionServicios.FacturacionProductos.RecordCount = 0 Then
'           MsgBox "Tiene que LEER el Estado de Cuenta del Paciente", vbInformation, "Resultado"
'           Exit Sub
'        End If
'        MousePointer = 11
'
'        'Crea nueva hoja
'        Set oExcel = GalenhosExcelApplication()  'New Excel.Application
'        Set oWorkBook = oExcel.Workbooks.Add
'        'Abre, copia y cierra la plantilla
'        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\ELiquidaCuenta.xls")
'        oWorkBookPlantilla.Worksheets("Liquidacion").Copy Before:=oWorkBook.Sheets(1)
'        oWorkBookPlantilla.Close
'        'Activa la primera hoja
'        Set oWorkSheet = oWorkBook.Sheets(1)
'        oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\Imagen de reportes.jpg"
'        'Inicio de Impresion
'
'        If ml_IdTipoServicio = sghConsultaExterna Then
'           Set rsReporte = mo_ReglasFacturacion.AtencionesFiltraIngresosPorIdAtencion(ml_idAtencion)
'        Else
'           Set rsReporte = mo_ReglasFacturacion.AtencionesFiltraEgresosPorIdAtencion(ml_idAtencion)
'        End If
'        If rsReporte.RecordCount = 0 Then
'           MousePointer = 1
'           Exit Sub
'        End If
'        Dim FuenteFinanciamiento As Long
'        FuenteFinanciamiento = rsReporte!idFuenteFinanciamiento
'
'
'        lcCabeceraLiquidacion = "'" & Trim(SIGHEntidades.RetornaNombrePC) & "/" & Trim(lcBuscaParametro.RetornaLoginUsuario(SIGHEntidades.Usuario))
'
'        oWorkSheet.Cells(1, 2).Value = "Estado de Cuenta al " & lcBuscaParametro.RetornaFechaHoraServidorSQL
'        oWorkSheet.Cells(2, 9).Value = lcCabeceraLiquidacion
'        oWorkSheet.Cells(3, 2).Value = "Cuenta:  " & Trim(Str(ml_idCuentaAtencion)) & "  " & txtEstadoCuenta.Text
'        oWorkSheet.Cells(3, 6).Value = IIf(txtDxEgr.Text = "", "", "Dx Egreso: " & txtDxEgr.Text)
'        oWorkSheet.Cells(4, 2).Value = "Paciente: " & txtPaciente.Text
'        oWorkSheet.Cells(4, 6).Value = "Nº Historia Clínica: " & Trim(txtNroHistoria.Text) & "       Dom.Pac: " & Trim(txtDomicilioPacienteEnAtencion.Text)
'        oWorkSheet.Cells(5, 2).Value = "F.Ingreso: " & rsReporte.Fields!FechaIngreso & " " & rsReporte.Fields!HoraIngreso
'        oWorkSheet.Cells(5, 6).Value = "Servicio Egreso: " & IIf(IsNull(rsReporte.Fields!codServicio), "", rsReporte.Fields!codServicio & " - " & rsReporte.Fields!DServicio)
'        oWorkSheet.Cells(6, 2).Value = "F.Alta Médica: " & IIf(IsNull(rsReporte.Fields!FechaEgreso), "", Format(rsReporte.Fields!FechaEgreso & " " & rsReporte.Fields!HoraEgreso, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HM))
'        oWorkSheet.Cells(6, 6).Value = "Cama: " & IIf(IsNull(rsReporte.Fields!codCama), "", rsReporte.Fields!codCama) & "     (Tarifa: " & lcdTipoFinanciamiento & ")" & "    " & ml_dCondicionAlta
'        iFila = 8
'        iCol = 2
'        ms_EstadosFacturacion = ""
'        ms_TiposFinanciamiento = ""
'        ml_AgruparPor = 1
'        lnTotal = 0: lnTotalSIS = 0: lnTotalSOAT = 0: lnTotalEXO = 0: lnTotalPAGO = 0: lnTotalDEBE = 0: lnTotalSALDO = 0: lnTotalCredito = 0
'
'        lnTDebe = 0: lnTPago = 0: lnTSaldo = 0
'
'        'Farmacia
'        Set rsReporte = ucFacturacionBienes.FacturacionProductos
'        If rsReporte.RecordCount > 0 Then
'            rsReporte.MoveFirst
'            oWorkSheet.Cells(iFila, 2).Value = "Farmacia"
'            lnTSubTotal = 0: lnTSubTotalSIS = 0: lnTSubTotalSOAT = 0: lnTSubTotalEXO = 0: lnTsubTotalConvenio = 0
'            lnTSubTotalPAGO = 0: lnTSubTotalDEBE = 0: lnTSubTotalSALDO = 0
'            Do While Not rsReporte.EOF
'              If rsReporte.Fields!IdOrden = 37 Or rsReporte.Fields!IdOrden = 38 Then
'               lnSIS = 0
'               End If
'               If rsReporte.Fields!IdEstadoFacturacion = 4 Or rsReporte.Fields!IdEstadoFacturacion = 1 Then   'Solo PAGADOS y REGISTRADOS
'                    lnSIS = IIf(IsNull(rsReporte.Fields!ImporteSIS), 0, rsReporte.Fields!ImporteSIS)
'                    lnEXO = IIf(IsNull(rsReporte.Fields!importeEXO), 0, rsReporte.Fields!importeEXO)
'                    lnSOAT = IIf(IsNull(rsReporte.Fields!ImporteSOAT), 0, rsReporte.Fields!ImporteSOAT)
'                    lnConvenio = IIf(IsNull(rsReporte.Fields!ImporteConv), 0, rsReporte.Fields!ImporteConv)
'                    lnCantidadPagarBienes = rsReporte.Fields("CantidadPagar").Value
'                    lnTotalPagarBienes = rsReporte.Fields("TotalPagar").Value
'                    If lbGeneraReciboPago = True Then
'                       lnDebe = lnTotalPagarBienes - lnEXO - lnSIS - lnSOAT
'                    Else
'                       If (rsReporte.Fields!CantidadSIS + rsReporte.Fields!CantidadSOAT + rsReporte.Fields!cantidadConv) = 0 Then
'                           lnDebe = lnTotalPagarBienes
'                       Else
'                           lnDebe = 0
'                       End If
'                    End If
'                    If rsReporte.Fields!IdEstadoFacturacion = 4 Then
'                       lnPago = rsReporte.Fields("ImporteEnBoleta").Value - lnEXO
'                    Else
'                       lnPago = 0
'                    End If
'
'                    lnSaldo = lnDebe - lnPago
'
'                    lnTSubTotal = lnTSubTotal + lnTotalPagarBienes
'                    lnTSubTotalEXO = lnTSubTotalEXO + lnEXO
'                    lnTSubTotalSIS = lnTSubTotalSIS + lnSIS
'                    lnTSubTotalSOAT = lnTSubTotalSOAT + lnSOAT
'                    lnTsubTotalConvenio = lnTsubTotalConvenio + lnConvenio
'                    lnTSubTotalPAGO = lnTSubTotalPAGO + lnPago
'                    lnTSubTotalDEBE = lnTSubTotalDEBE + lnDebe
'                    lnTSubTotalSALDO = lnTSubTotalSALDO + lnSaldo
'
'                    lnTotal = lnTotal + lnTotalPagarBienes
'                    lnTotalSIS = lnTotalSIS + lnSIS
'                    lnTotalSOAT = lnTotalSOAT + lnSOAT
'                    lnTotalEXO = lnTotalEXO + lnEXO
'                    lnTotalConvenio = lnTotalConvenio + lnConvenio
'                    lnTotalPAGO = lnTotalPAGO + lnPago
'                    lnTotalDEBE = lnTotalDEBE + lnDebe
'                    lnTotalSALDO = lnTotalSALDO + lnSaldo
'
'                    If rsReporte.Fields!idProducto = lnIdPagosACuenta Then   'Pagos a cuenta
'                       lnTotalCredito = lnTotalCredito + lnTotalPagarBienes
'                    End If
'
'                End If
'                rsReporte.MoveNext
'             Loop
'             mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 5, iFila, iCol + 8
'             If FuenteFinanciamiento = 1 Or FuenteFinanciamiento = 5 Then
'               oWorkSheet.Cells(iFila, 5).Value = Format(lnTSubTotal, "0.00")
'             Else
'               oWorkSheet.Cells(iFila, 5).Value = Format(0, "0.00")
'             End If
'             oWorkSheet.Cells(iFila, 6).Value = Format(lnTSubTotalEXO, "0.00")
'             oWorkSheet.Cells(iFila, 7).Value = Format(lnTSubTotalSIS, "0.00")
'             oWorkSheet.Cells(iFila, 8).Value = Format(lnTSubTotalSOAT, "0.00")
'             oWorkSheet.Cells(iFila, 9).Value = Format(lnTsubTotalConvenio, "0.00")
'             oWorkSheet.Cells(iFila, 10).Value = Format(lnTSubTotalDEBE, "0.00")
'             iFila = iFila + 1
'        End If
'        rsReporte.Close
'        lnPagoEnFarmacia = lnTSubTotalSALDO
'        'Servicios
'        Set rsReporte = ucFacturacionServicios.FacturacionProductos
'        'debb-jamo
'        If txtFegreso.Text = "" And ml_IdTipoServicio = sghHospitalizacion Then
'           ldFechaAlta = CDate(Format(Now, SIGHEntidades.DevuelveFechaSoloFormato_DMY))
'           lcHoraAlta = Format(Now, SIGHEntidades.DevuelveHoraSoloFormato_HM)
'           mo_AdminAdmision.GeneraEstanciaPorCadaServicioHospitalizado ml_idCuentaAtencion, ldFechaAlta, lcHoraAlta, rsReporte, lnTotalPagarEstancia, lnTotalDiasEstancia, oConexion
'           Select Case mo_ReglasFacturacion.TiposFinanciamientoDevuelveComoSeTrabajaEnEstadoCuenta(lnIdTipoFinanciamientoActual, oConexion)
'           Case sghTrabajaSeguroSIS
'               lnTotalSIS = lnTotalSIS
'           Case sghTrabajaSeguroSOAT
'               lnTotalSOAT = lnTotalSOAT
'           Case sghTrabajaSeguroConvenios
'               lnTotalConvenio = lnTotalConvenio
'           Case Else
'               txtTotalApagar.Text = Val(txtTotalApagar.Text) + lnTotalPagarEstancia
'           End Select
'        End If
'        lnTotalPagosAdelantados = 0
'        If rsReporte.RecordCount > 0 Then
'            rsReporte.Sort = "IdPuntoCarga"
'            rsReporte.MoveFirst
'            Do While Not rsReporte.EOF
'
'                idPuntoCarga = rsReporte.Fields("IdPuntoCarga").Value
'                oWorkSheet.Cells(iFila, 2).Value = mo_reglasComunes.FactPuntosCargaSeleccionarPorIdDevDescripcion(rsReporte.Fields("IdPuntoCarga").Value)
'
'                lnTSubTotal = 0: lnTSubTotalSIS = 0: lnTSubTotalSOAT = 0: lnTSubTotalEXO = 0: lnTsubTotalConvenio = 0
'                lnTSubTotalPAGO = 0: lnTSubTotalDEBE = 0: lnTSubTotalSALDO = 0
'                Do While Not rsReporte.EOF And idPuntoCarga = rsReporte.Fields("IdPuntoCarga").Value
'                 If rsReporte.Fields!IdOrden = 641 Then
'                  lnSIS = 0
'                 End If
'                    If rsReporte.Fields!IdEstadoFacturacion = 4 Or rsReporte.Fields!IdEstadoFacturacion = 1 Or rsReporte.Fields!IdEstadoFacturacion = 10 Or rsReporte.Fields!IdEstadoFacturacion = sghConPreVenta Then  'Solo PAGADOS/REGISTRADOS/AUTORIZ.AUTOMATICA/preventa
'                        lnSIS = IIf(IsNull(rsReporte.Fields!ImporteSIS), 0, rsReporte.Fields!ImporteSIS)
'                        lnSOAT = IIf(IsNull(rsReporte.Fields!ImporteSOAT), 0, rsReporte.Fields!ImporteSOAT)
'                        lnEXO = IIf(IsNull(rsReporte.Fields!importeEXO), 0, rsReporte.Fields!importeEXO)
'                        lnConvenio = IIf(IsNull(rsReporte.Fields!ImporteConv), 0, rsReporte.Fields!ImporteConv)
'                        If rsReporte.Fields!IdEstadoFacturacion = 4 Then
'                            lnPago = rsReporte.Fields("ImporteEnBoleta").Value - lnEXO
'                        Else
'                            lnPago = 0
'                        End If
'                        If rsReporte.Fields!idProducto <> lnIdPagosACuenta And rsReporte.Fields!idProducto <> lnIdPagosXdevoluciones Then
'                           If lbGeneraReciboPago = True Then
'                                lnDebe = rsReporte.Fields!TotalPagar - lnEXO - lnSIS - lnSOAT
'                           Else
'                                If (rsReporte.Fields!CantidadSIS + rsReporte.Fields!CantidadSOAT + rsReporte.Fields!cantidadConv) = 0 Then
'                                    lnDebe = rsReporte.Fields!TotalPagar
'                                Else
'                                    lnDebe = 0
'                                End If
'                           End If
'                           lnSaldo = lnDebe - lnPago
'                        Else
'                           If rsReporte.Fields!idProducto = lnIdPagosACuenta Then
'                                lnTotalPagosAdelantados = lnTotalPagosAdelantados + rsReporte.Fields!ImporteEnBoleta
'                                lnDebe = 0
'                                If lbGeneraReciboPago = True Then
'                                   lnSaldo = -rsReporte.Fields!ImporteEnBoleta
'                                Else
'                                   lnSaldo = 0
'                                End If
'                           Else
'                                lnDebe = 0
'                                If lbGeneraReciboPago = True Then
'                                   lnSaldo = rsReporte.Fields!ImporteEnBoleta
'                                   lnPago = -rsReporte.Fields!ImporteEnBoleta
'                                Else
'                                   lnSaldo = 0
'                                   lnPago = 0
'                                End If
'                           End If
'                        End If
'
'                        If rsReporte.Fields!idProducto <> lnIdPagosACuenta Then
'                           lnTSubTotal = lnTSubTotal + rsReporte.Fields("TotalPagar").Value
'                           lnTSubTotalDEBE = lnTSubTotalDEBE + lnDebe
'                        End If
'                        lnTSubTotalEXO = lnTSubTotalEXO + lnEXO
'                        lnTSubTotalSIS = lnTSubTotalSIS + lnSIS
'                        lnTSubTotalSOAT = lnTSubTotalSOAT + lnSOAT
'                        lnTsubTotalConvenio = lnTsubTotalConvenio + lnConvenio
'                        lnTSubTotalPAGO = lnTSubTotalPAGO + lnPago
'                        lnTSubTotalSALDO = lnTSubTotalSALDO + lnSaldo
'
'                        If rsReporte.Fields!idProducto <> lnIdPagosACuenta Then
'                           lnTotal = lnTotal + rsReporte.Fields("TotalPagar").Value
'                           lnTotalDEBE = lnTotalDEBE + lnDebe
'                        End If
'                        lnTotalSIS = lnTotalSIS + lnSIS
'                        lnTotalSOAT = lnTotalSOAT + lnSOAT
'                        lnTotalEXO = lnTotalEXO + lnEXO
'                        lnTotalConvenio = lnTotalConvenio + lnConvenio
'                        lnTotalPAGO = lnTotalPAGO + lnPago
'                        lnTotalSALDO = lnTotalSALDO + lnSaldo
'
'                        If rsReporte.Fields!idProducto = lnIdPagosACuenta Then   'Pagos a cuenta
'                           lnTotalCredito = lnTotalCredito + rsReporte.Fields("TotalPorPagar").Value
'                        End If
'
'                    End If
'                    rsReporte.MoveNext
'                    If rsReporte.EOF Then
'                       Exit Do
'                    End If
'                Loop
'                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 5, iFila, iCol + 8
'                If FuenteFinanciamiento = 1 Or FuenteFinanciamiento = 5 Then
'                  oWorkSheet.Cells(iFila, 5).Value = Format(lnTSubTotal, "0.00")
'                Else
'                  oWorkSheet.Cells(iFila, 5).Value = Format(0, "0.00")
'                End If
'                oWorkSheet.Cells(iFila, 6).Value = Format(lnTSubTotalEXO, "0.00")
'                oWorkSheet.Cells(iFila, 7).Value = Format(lnTSubTotalSIS, "0.00")
'                oWorkSheet.Cells(iFila, 8).Value = Format(lnTSubTotalSOAT, "0.00")
'                oWorkSheet.Cells(iFila, 9).Value = Format(lnTsubTotalConvenio, "0.00")
'                oWorkSheet.Cells(iFila, 10).Value = Format(lnTSubTotalDEBE, "0.00")
'                iFila = iFila + 1
'             Loop
'        End If
'        iFila = iFila + 1
'        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, iCol + 8
'        oWorkSheet.Cells(iFila, 2).Value = "Total: "
'        If FuenteFinanciamiento = 1 Or FuenteFinanciamiento = 5 Then
'          oWorkSheet.Cells(iFila, 5).Value = Format(lnTotal, "0.00")
'        Else
'          oWorkSheet.Cells(iFila, 5).Value = Format(0, "0.00")
'        End If
'        oWorkSheet.Cells(iFila, 6).Value = Format(lnTotalEXO, "0.00")
'        oWorkSheet.Cells(iFila, 7).Value = Format(lnTotalSIS, "0.00")
'        oWorkSheet.Cells(iFila, 8).Value = Format(lnTotalSOAT, "0.00")
'        oWorkSheet.Cells(iFila, 9).Value = Format(lnTotalConvenio, "0.00")
'        oWorkSheet.Cells(iFila, 10).Value = Format(lnTotalDEBE, "0.00")
'        If lbGeneraReciboPago = True Then
'            If lnTotalPagosAdelantados > lnPagoEnFarmacia Then
'               lnTotalPagosAdelantados = lnTotalPagosAdelantados - lnPagoEnFarmacia
'               lnPagoEnFarmacia = 0
'            Else
'               lnPagoEnFarmacia = lnPagoEnFarmacia - lnTotalPagosAdelantados
'               lnTotalPagosAdelantados = 0
'            End If
'            lnPagoEnServicios = lnPagoEnServicios - lnTotalPagosAdelantados
'        Else
'           Select Case mo_ReglasFacturacion.TiposFinanciamientoDevuelveComoSeTrabajaEnEstadoCuenta(lnIdTipoFinanciamientoActual, oConexion)
'           Case sghTrabajaSeguroSIS
'               lnTotalSIS = lnTotalSIS - lnTotalPagosAdelantados
'           Case sghTrabajaSeguroSOAT
'               lnTotalSOAT = lnTotalSOAT - lnTotalPagosAdelantados
'           Case sghTrabajaSeguroConvenios
'               lnTotalConvenio = lnTotalConvenio - lnTotalPagosAdelantados
'           End Select
'        End If
'        lnPagoEnServicios = lnTotalSALDO - lnPagoEnFarmacia
'
'        iFila = iFila + 2
'        oWorkSheet.Cells(iFila, 2).Value = "TOTAL PAGO A CUENTA: "
'        oWorkSheet.Cells(iFila, 4).Value = Format(lnTotalPagosAdelantados, "0.00")
'        iFila = iFila + 1
'        oWorkSheet.Cells(iFila, 2).Value = "TOTAL PACIENTE: "
'        oWorkSheet.Cells(iFila, 4).Value = Format(lnTotalDEBE, "0.00") '
'        iFila = iFila + 1
'
'        If lbGeneraReciboPago = True Then
'            oWorkSheet.Cells(iFila, 2).Value = "PAGADO: "
'            oWorkSheet.Cells(iFila, 4).Value = Format(lnTotalPAGO - lnTotalPagosAdelantados + lnPagosXdevoluciones, "0.00") 'oWorkSheet.Cells(iFila, 4).Value = Format((lnTotalDEBE - Adelanto - lnTotalPAGO), "0.00")
'            iFila = iFila + 1
'            iFila = iFila + 1
'            iFila = iFila + 1
'            oWorkSheet.Cells(iFila, 2).Value = "POR PAGAR PACIENTE: "
'            oWorkSheet.Cells(iFila, 4).Value = Format(Val(txtTotalApagar.Text), "0.00")  'oWorkSheet.Cells(iFila, 4).Value = Format(lnTotalSALDO + lnTotalPagosAdelantados, "0.00")
'        Else
'            oWorkSheet.Cells(iFila, 2).Value = "PAGADO: "
'            oWorkSheet.Cells(iFila, 4).Value = Format(lnTotalPAGO - lnTotalPagosAdelantados, "0.00") 'oWorkSheet.Cells(iFila, 4).Value = Format((lnTotalDEBE - Adelanto - lnTotalPAGO), "0.00")
'            iFila = iFila + 1
'            iFila = iFila + 1
'            iFila = iFila + 1
'            lnTotalSALDO = Val(txtTotalServicios.Text) + Val(txtTotalFarmacia.Text)
'            If lnTotalSALDO > 0 Then
'                oWorkSheet.Cells(iFila, 2).Value = "POR PAGAR PACIENTE: "
'                oWorkSheet.Cells(iFila, 4).Value = Format(lnTotalSALDO, "0.00") 'oWorkSheet.Cells(iFila, 4).Value = Format(lnTotalSALDO + lnTotalPagosAdelantados, "0.00")
'                iFila = iFila + 1
'            End If
'        End If
'        '
'        If lnTotalSOAT > 0 Then
'          oWorkSheet.Cells(iFila, 2).Value = "TOTAL A PAGAR SOAT: "
'          oWorkSheet.Cells(iFila, 4).Value = Format(lnTotalSOAT, "0.00")
'        ElseIf lnTotalSIS > 0 Then
'          oWorkSheet.Cells(iFila, 2).Value = "TOTAL A PAGAR SIS: "
'          oWorkSheet.Cells(iFila, 4).Value = Format(lnTotalSIS, "0.00")
'        ElseIf lnTotalConvenio > 0 Then
'          oWorkSheet.Cells(iFila, 2).Value = "TOTAL A PAGAR CONVENIO: "
'          oWorkSheet.Cells(iFila, 4).Value = Format(lnTotalConvenio, "0.00")
'        End If
'        '
'        If lcListaDeOrdenesDePago <> "" Then
'            iFila = iFila + 2
'            oWorkSheet.Cells(iFila, 2).Value = "* El CAJERO debe emitir Boletas usando " & lcListaDeOrdenesDePago
'        End If
'        '***Donaciones en Farmacia
'        If rsItemsDonaciones.RecordCount > 0 Then
'           Dim lnCantidadDona As Long, lcCodigoDona As String, lbContinua As Boolean
'           iFila = iFila + 2
'           oWorkSheet.Cells(iFila, 2).Value = "LISTA DE DONACIONES:"
'           iFila = iFila + 1
'           oWorkSheet.Cells(iFila, 2).Value = "Descripción"
'           oWorkSheet.Cells(iFila, 9).Value = "Cantidad"
'           mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 9
'           Set rsReporte = Nothing
'           With rsReporte
'                  .Fields.Append "Codigo", adVarChar, 10, adFldIsNullable
'                  .Fields.Append "Nombre", adVarChar, 200, adFldIsNullable
'                  .Fields.Append "Cantidad", adInteger
'                  .CursorType = adOpenDynamic
'                  .LockType = adLockOptimistic
'                  .Open
'           End With
'           rsItemsDonaciones.MoveFirst
'           Do While Not rsItemsDonaciones.EOF
'              lbContinua = True
'              If rsReporte.RecordCount > 0 Then
'                 rsReporte.MoveFirst
'                 rsReporte.Find "Codigo='" & rsItemsDonaciones.Fields!Codigo & "'"
'                 If Not rsReporte.EOF Then
'                    lbContinua = False
'                 End If
'              End If
'              If lbContinua = True Then
'                  rsReporte.AddNew
'                  rsReporte.Fields!Codigo = rsItemsDonaciones.Fields!Codigo
'                  rsReporte.Fields!Nombre = rsItemsDonaciones.Fields!Nombre
'              End If
'              rsReporte.Fields!Cantidad = rsReporte.Fields!Cantidad + rsItemsDonaciones.Fields!Cantidad
'              rsReporte.Update
'              rsItemsDonaciones.MoveNext
'           Loop
'           rsReporte.Sort = "nombre,codigo"
'           rsReporte.MoveFirst
'           Do While Not rsReporte.EOF
'              iFila = iFila + 1
'              lcCodigoDona = rsReporte.Fields!Codigo
'              oWorkSheet.Cells(iFila, 2).Value = Trim(rsReporte.Fields!Codigo) & " " & rsReporte.Fields!Nombre
'              oWorkSheet.Cells(iFila, 9).Value = rsReporte.Fields!Cantidad
'              rsReporte.MoveNext
'           Loop
'        End If
'        '
'        oWorkSheet.PageSetup.PrintTitleRows = "$1:$7"
'        If oWorkSheet.PageSetup.PrintArea <> "" Then oWorkSheet.PageSetup.PrintArea = "$A$1:$K$" & (iFila + 2) 'sighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
'
'
'
'            oExcel.Visible = True
'
'            oWorkSheet.PrintPreview
'
'    End If
'    MousePointer = 1
'    'debb-jamo
'    If txtFegreso.Text = "" And ml_IdTipoServicio = sghHospitalizacion Then
'       LimparDatos
'    End If
'    oConexion.Close
'    Set oConexion = Nothing
'    Set rsReporte = Nothing
'    Set oWorkSheet = Nothing
'    Set oExcel = Nothing
'    '
'    mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar ml_idCuentaAtencion, False, 0
    Dim iFila As Long, iCol As Integer
    Dim rsReporte As New Recordset
    Dim ms_EstadosFacturacion As String
    Dim ms_TiposFinanciamiento As String
    Dim ml_AgruparPor As Long
    Dim mo_ReporteUtil As New sighentidades.ReporteUtil
    Dim idPuntoCarga As Long
    
    Dim lnTSubTotal As Double: Dim lnTSubTotalSIS As Double: Dim lnTSubTotalSOAT As Double: Dim lnTSubTotalEXO As Double: Dim lnTsubTotalConvenio As Double
    Dim lnTSubTotalPAGO As Double: Dim lnTSubTotalDEBE As Double: Dim lnTSubTotalSALDO As Double
    
    Dim lnTotal As Double: Dim lnTotalSIS As Double: Dim lnTotalSOAT As Double: Dim lnTotalEXO As Double: Dim lnTotalConvenio As Double
    Dim lnTotalPAGO As Double: Dim lnTotalDEBE As Double: Dim lnTotalSALDO As Double
    
    Dim lnDebe As Double: Dim lnPago As Double: Dim lnSaldo As Double
    Dim lnTDebe As Double: Dim lnTPago As Double: Dim lnTSaldo As Double
    Dim lnSIS As Double: Dim lnSOAT As Double: Dim lnEXO As Double: Dim lnTotalCredito As Double: Dim lnConvenio As Double
    Dim lnDctos As Double: Dim lnPagoEnFarmacia As Double: Dim lnPagoEnServicios As Double
    Dim lnTotalPagosAdelantados As Double
    Dim lnCantidadPagarBienes As Long, lnTotalPagarBienes As Double
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim lnTotalPagarEstancia As Double, lnTotalDiasEstancia As Long
    Dim ldFechaAlta As Date, lcHoraAlta As String, lcCabeceraLiquidacion As String
    Dim oConexion As New Connection
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
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
        Dim ml_lnHWnd As Long
        Dim lnHWnd As Long
    Else
        Dim oExcel As Excel.Application
        Dim oWorkBookPlantilla As Workbook
        Dim oWorkBook As Workbook
        Dim oWorkSheet As Worksheet
        Dim oRange As range
        Dim range As Excel.range
        Dim borders As Excel.borders
    End If
    
    If txtPaciente.Text = "" Then
        MsgBox "Tiene que LEER el Estado de Cuenta del Paciente", vbInformation, "Resultado"
    Else
        If ucFacturacionBienes.FacturacionProductos.RecordCount = 0 And ucFacturacionServicios.FacturacionProductos.RecordCount = 0 Then
           MsgBox "Tiene que LEER el Estado de Cuenta del Paciente", vbInformation, "Resultado"
           Exit Sub
        End If
        MousePointer = 11
        If lbEsOpenOffice = True Then
            'Abre el archivo ExcelOpenOffice
            lcArchivoExcel = App.Path + "\Plantillas\ELiquidaCuenta.ods"
    '        FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
    '        Chemin = "file:///" & App.Path & "\Plantillas\"
    '        Chemin = Replace(Chemin, "\", "/")
    '        Fichier = Chemin & "/OpenOffice.ods"
            '
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
            'Encabezado de Pagina
            mo_CabeceraReportes.CabeceraReportes Document, True
            ' Pone la ventana en primer plano, pasándole el Hwnd
            ret = SetForegroundWindow(lnHWnd)
        Else
            'Crea nueva hoja
            Set oExcel = GalenhosExcelApplication()  'New Excel.Application
            Set oWorkBook = oExcel.Workbooks.Add
            'Abre, copia y cierra la plantilla
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\ELiquidaCuenta.xls")
            oWorkBookPlantilla.Worksheets("Liquidacion").Copy Before:=oWorkBook.Sheets(1)
            oWorkBookPlantilla.Close
            'Activa la primera hoja
            Set oWorkSheet = oWorkBook.Sheets(1)
            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
            'Inicio de Impresion
        End If
        If ml_IdTipoServicio = sghConsultaExterna Then
           Set rsReporte = mo_ReglasFacturacion.AtencionesFiltraIngresosPorIdAtencion(ml_idAtencion)
        Else
           Set rsReporte = mo_ReglasFacturacion.AtencionesFiltraEgresosPorIdAtencion(ml_idAtencion)
        End If
        If rsReporte.RecordCount = 0 Then
           MousePointer = 1
           Exit Sub
        End If
        Dim FuenteFinanciamiento As Long
        FuenteFinanciamiento = rsReporte!IdFuenteFinanciamiento
        lcCabeceraLiquidacion = "'" & Trim(sighentidades.RetornaNombrePC) & "/" & Trim(lcBuscaParametro.RetornaLoginUsuario(sighentidades.Usuario))
        
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(1, 0).setFormula("Estado de Cuenta al " & lcBuscaParametro.RetornaFechaHoraServidorSQL)
            Call Feuille.getcellbyposition(8, 1).setFormula(lcCabeceraLiquidacion)
            Call Feuille.getcellbyposition(1, 2).setFormula("Cuenta:  " & Trim(Str(ml_idCuentaAtencion)) & "  " & txtEstadoCuenta.Text)
            Call Feuille.getcellbyposition(5, 2).setFormula(IIf(txtDxEgr.Text = "", "", "Dx Egreso: " & txtDxEgr.Text))
            Call Feuille.getcellbyposition(1, 3).setFormula("Paciente: " & txtPaciente.Text)
            Call Feuille.getcellbyposition(5, 3).setFormula("Nº Historia Clínica: " & Trim(txtNroHistoria.Text) & "       Dom.Pac: " & Trim(txtDomicilioPacienteEnAtencion.Text))
            Call Feuille.getcellbyposition(1, 4).setFormula("F.Ingreso: " & rsReporte.Fields!FechaIngreso & " " & rsReporte.Fields!HoraIngreso)
            Call Feuille.getcellbyposition(5, 4).setFormula("Servicio Egreso: " & IIf(IsNull(rsReporte.Fields!codServicio), "", rsReporte.Fields!codServicio & " - " & rsReporte.Fields!DServicio))
            Call Feuille.getcellbyposition(1, 5).setFormula("F.Alta Médica: " & IIf(IsNull(rsReporte.Fields!FechaEgreso), "", Format(rsReporte.Fields!FechaEgreso & " " & rsReporte.Fields!HoraEgreso, sighentidades.DevuelveFechaSoloFormato_DMY_HM)))
            Call Feuille.getcellbyposition(5, 5).setFormula("Cama: " & IIf(IsNull(rsReporte.Fields!codCama), "", rsReporte.Fields!codCama) & "     (Tarifa: " & lcdTipoFinanciamiento & ")" & "    " & ml_dCondicionAlta)
        Else
            oWorkSheet.Cells(1, 2).Value = "Estado de Cuenta al " & lcBuscaParametro.RetornaFechaHoraServidorSQL
            oWorkSheet.Cells(2, 9).Value = lcCabeceraLiquidacion
            oWorkSheet.Cells(3, 2).Value = "Cuenta:  " & Trim(Str(ml_idCuentaAtencion)) & "  " & txtEstadoCuenta.Text
            oWorkSheet.Cells(3, 6).Value = IIf(txtDxEgr.Text = "", "", "Dx Egreso: " & txtDxEgr.Text)
            oWorkSheet.Cells(4, 2).Value = "Paciente: " & txtPaciente.Text
            oWorkSheet.Cells(4, 6).Value = "Nº Historia Clínica: " & Trim(txtNroHistoria.Text) & "       Dom.Pac: " & Trim(txtDomicilioPacienteEnAtencion.Text)
            oWorkSheet.Cells(5, 2).Value = "F.Ingreso: " & rsReporte.Fields!FechaIngreso & " " & rsReporte.Fields!HoraIngreso
            oWorkSheet.Cells(5, 6).Value = "Servicio Egreso: " & IIf(IsNull(rsReporte.Fields!codServicio), "", rsReporte.Fields!codServicio & " - " & rsReporte.Fields!DServicio)
            oWorkSheet.Cells(6, 2).Value = "F.Alta Médica: " & IIf(IsNull(rsReporte.Fields!FechaEgreso), "", Format(rsReporte.Fields!FechaEgreso & " " & rsReporte.Fields!HoraEgreso, sighentidades.DevuelveFechaSoloFormato_DMY_HM))
            oWorkSheet.Cells(6, 6).Value = "Cama: " & IIf(IsNull(rsReporte.Fields!codCama), "", rsReporte.Fields!codCama) & "     (Tarifa: " & lcdTipoFinanciamiento & ")" & "    " & ml_dCondicionAlta
        End If
        iFila = 8
        iCol = 2
        ms_EstadosFacturacion = ""
        ms_TiposFinanciamiento = ""
        ml_AgruparPor = 1
        lnTotal = 0: lnTotalSIS = 0: lnTotalSOAT = 0: lnTotalEXO = 0: lnTotalPAGO = 0: lnTotalDEBE = 0: lnTotalSALDO = 0: lnTotalCredito = 0
        
        lnTDebe = 0: lnTPago = 0: lnTSaldo = 0
        
        'Farmacia
        Set rsReporte = ucFacturacionBienes.FacturacionProductos
        If rsReporte.RecordCount > 0 Then
            rsReporte.MoveFirst
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Farmacia")
            Else
                oWorkSheet.Cells(iFila, 2).Value = "Farmacia"
            End If
            lnTSubTotal = 0: lnTSubTotalSIS = 0: lnTSubTotalSOAT = 0: lnTSubTotalEXO = 0: lnTsubTotalConvenio = 0
            lnTSubTotalPAGO = 0: lnTSubTotalDEBE = 0: lnTSubTotalSALDO = 0
            Do While Not rsReporte.EOF
              If rsReporte.Fields!IdOrden = 37 Or rsReporte.Fields!IdOrden = 38 Then
               lnSIS = 0
               End If
               If rsReporte.Fields!idestadofacturacion = 4 Or rsReporte.Fields!idestadofacturacion = 1 Then   'Solo PAGADOS y REGISTRADOS
                    lnSIS = IIf(IsNull(rsReporte.Fields!ImporteSIS), 0, rsReporte.Fields!ImporteSIS)
                    lnEXO = IIf(IsNull(rsReporte.Fields!importeEXO), 0, rsReporte.Fields!importeEXO)
                    lnSOAT = IIf(IsNull(rsReporte.Fields!ImporteSOAT), 0, rsReporte.Fields!ImporteSOAT)
                    lnConvenio = IIf(IsNull(rsReporte.Fields!ImporteConv), 0, rsReporte.Fields!ImporteConv)
                    lnCantidadPagarBienes = rsReporte.Fields("CantidadPagar").Value
                    lnTotalPagarBienes = rsReporte.Fields("TotalPagar").Value
                    If lbGeneraReciboPago = True Then
                       lnDebe = lnTotalPagarBienes - lnEXO - lnSIS - lnSOAT
                    Else
                       If (rsReporte.Fields!CantidadSIS + rsReporte.Fields!CantidadSOAT + rsReporte.Fields!cantidadConv) = 0 Then
                           lnDebe = lnTotalPagarBienes
                       Else
                           lnDebe = 0
                       End If
                    End If
                    If rsReporte.Fields!idestadofacturacion = 4 Then
                       lnPago = rsReporte.Fields("ImporteEnBoleta").Value - lnEXO
                    Else
                       lnPago = 0
                    End If

                    lnSaldo = lnDebe - lnPago
                    
                    lnTSubTotal = lnTSubTotal + lnTotalPagarBienes
                    lnTSubTotalEXO = lnTSubTotalEXO + lnEXO
                    lnTSubTotalSIS = lnTSubTotalSIS + lnSIS
                    lnTSubTotalSOAT = lnTSubTotalSOAT + lnSOAT
                    lnTsubTotalConvenio = lnTsubTotalConvenio + lnConvenio
                    lnTSubTotalPAGO = lnTSubTotalPAGO + lnPago
                    lnTSubTotalDEBE = lnTSubTotalDEBE + lnDebe
                    lnTSubTotalSALDO = lnTSubTotalSALDO + lnSaldo
                    
                    lnTotal = lnTotal + lnTotalPagarBienes
                    lnTotalSIS = lnTotalSIS + lnSIS
                    lnTotalSOAT = lnTotalSOAT + lnSOAT
                    lnTotalEXO = lnTotalEXO + lnEXO
                    lnTotalConvenio = lnTotalConvenio + lnConvenio
                    lnTotalPAGO = lnTotalPAGO + lnPago
                    lnTotalDEBE = lnTotalDEBE + lnDebe
                    lnTotalSALDO = lnTotalSALDO + lnSaldo
            
                    If rsReporte.Fields!idProducto = lnIdPagosACuenta Then   'Pagos a cuenta
                       lnTotalCredito = lnTotalCredito + lnTotalPagarBienes
                    End If
                    
                End If
                rsReporte.MoveNext
             Loop
             If lbEsOpenOffice = True Then
                Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(5) & CStr(iFila) & ":" & mo_ReglasReportes.BuscaNombreColumna(iCol + 8) & CStr(iFila))
                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
             Else
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 5, iFila, iCol + 8
             End If
             If lbEsOpenOffice = True Then
                If FuenteFinanciamiento = 1 Or FuenteFinanciamiento = 5 Then
                  Call Feuille.getcellbyposition(4, iFila - 1).setFormula(Format(lnTSubTotal, "0.00"))
                Else
                  Call Feuille.getcellbyposition(4, iFila - 1).setFormula(Format(0, "0.00"))
                End If
             Else
                If FuenteFinanciamiento = 1 Or FuenteFinanciamiento = 5 Then
                  oWorkSheet.Cells(iFila, 5).Value = Format(lnTSubTotal, "0.00")
                Else
                  oWorkSheet.Cells(iFila, 5).Value = Format(0, "0.00")
                End If
             End If
             If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(5, iFila - 1).setFormula(Format(lnTSubTotalEXO, "0.00"))
                Call Feuille.getcellbyposition(6, iFila - 1).setFormula(Format(lnTSubTotalSIS, "0.00"))
                Call Feuille.getcellbyposition(7, iFila - 1).setFormula(Format(lnTSubTotalSOAT, "0.00"))
                Call Feuille.getcellbyposition(8, iFila - 1).setFormula(Format(lnTsubTotalConvenio, "0.00"))
                Call Feuille.getcellbyposition(9, iFila - 1).setFormula(Format(lnTSubTotalDEBE, "0.00"))
             Else
                oWorkSheet.Cells(iFila, 6).Value = Format(lnTSubTotalEXO, "0.00")
                oWorkSheet.Cells(iFila, 7).Value = Format(lnTSubTotalSIS, "0.00")
                oWorkSheet.Cells(iFila, 8).Value = Format(lnTSubTotalSOAT, "0.00")
                oWorkSheet.Cells(iFila, 9).Value = Format(lnTsubTotalConvenio, "0.00")
                oWorkSheet.Cells(iFila, 10).Value = Format(lnTSubTotalDEBE, "0.00")
             End If
             iFila = iFila + 1
        End If
        rsReporte.Close
        lnPagoEnFarmacia = lnTSubTotalSALDO
        'Servicios
        Set rsReporte = ucFacturacionServicios.FacturacionProductos
        'debb-jamo
        If txtFegreso.Text = "" And ml_IdTipoServicio = sghHospitalizacion Then
           ldFechaAlta = CDate(Format(Now, sighentidades.DevuelveFechaSoloFormato_DMY))
           lcHoraAlta = Format(Now, sighentidades.DevuelveHoraSoloFormato_HM)
           mo_AdminAdmision.GeneraEstanciaPorCadaServicioHospitalizado ml_idCuentaAtencion, ldFechaAlta, lcHoraAlta, rsReporte, lnTotalPagarEstancia, lnTotalDiasEstancia, oConexion
           Select Case mo_ReglasFacturacion.TiposFinanciamientoDevuelveComoSeTrabajaEnEstadoCuenta(lnIdTipoFinanciamientoActual, oConexion)
           Case sghTrabajaSeguroSIS
               lnTotalSIS = lnTotalSIS
           Case sghTrabajaSeguroSOAT
               lnTotalSOAT = lnTotalSOAT
           Case sghTrabajaSeguroConvenios
               lnTotalConvenio = lnTotalConvenio
           Case Else
               txtTotalApagar.Text = Val(txtTotalApagar.Text) + lnTotalPagarEstancia
           End Select
        End If
        lnTotalPagosAdelantados = 0
        If rsReporte.RecordCount > 0 Then
            rsReporte.Sort = "IdPuntoCarga"
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF

                idPuntoCarga = rsReporte.Fields("IdPuntoCarga").Value
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(1, iFila - 1).setFormula(mo_ReglasComunes.FactPuntosCargaSeleccionarPorIdDevDescripcion(rsReporte.Fields("IdPuntoCarga").Value))
                Else
                    oWorkSheet.Cells(iFila, 2).Value = mo_ReglasComunes.FactPuntosCargaSeleccionarPorIdDevDescripcion(rsReporte.Fields("IdPuntoCarga").Value)
                End If
                lnTSubTotal = 0: lnTSubTotalSIS = 0: lnTSubTotalSOAT = 0: lnTSubTotalEXO = 0: lnTsubTotalConvenio = 0
                lnTSubTotalPAGO = 0: lnTSubTotalDEBE = 0: lnTSubTotalSALDO = 0
                Do While Not rsReporte.EOF And idPuntoCarga = rsReporte.Fields("IdPuntoCarga").Value
                 If rsReporte.Fields!IdOrden = 641 Then
                  lnSIS = 0
                 End If
                    If rsReporte.Fields!idestadofacturacion = 4 Or rsReporte.Fields!idestadofacturacion = 1 Or rsReporte.Fields!idestadofacturacion = 10 Or rsReporte.Fields!idestadofacturacion = sghConPreVenta Then  'Solo PAGADOS/REGISTRADOS/AUTORIZ.AUTOMATICA/preventa
                        lnSIS = IIf(IsNull(rsReporte.Fields!ImporteSIS), 0, rsReporte.Fields!ImporteSIS)
                        lnSOAT = IIf(IsNull(rsReporte.Fields!ImporteSOAT), 0, rsReporte.Fields!ImporteSOAT)
                        lnEXO = IIf(IsNull(rsReporte.Fields!importeEXO), 0, rsReporte.Fields!importeEXO)
                        lnConvenio = IIf(IsNull(rsReporte.Fields!ImporteConv), 0, rsReporte.Fields!ImporteConv)
                        If rsReporte.Fields!idestadofacturacion = 4 Then
                            lnPago = rsReporte.Fields("ImporteEnBoleta").Value - lnEXO
                        Else
                            lnPago = 0
                        End If
                        If rsReporte.Fields!idProducto <> lnIdPagosACuenta And rsReporte.Fields!idProducto <> lnIdPagosXdevoluciones Then
                           If lbGeneraReciboPago = True Then
                                lnDebe = rsReporte.Fields!TotalPagar - lnEXO - lnSIS - lnSOAT
                           Else
                                If (rsReporte.Fields!CantidadSIS + rsReporte.Fields!CantidadSOAT + rsReporte.Fields!cantidadConv) = 0 Then
                                    lnDebe = rsReporte.Fields!TotalPagar
                                Else
                                    lnDebe = 0
                                End If
                           End If
                           lnSaldo = lnDebe - lnPago
                        Else
                           If rsReporte.Fields!idProducto = lnIdPagosACuenta Then
                                lnTotalPagosAdelantados = lnTotalPagosAdelantados + rsReporte.Fields!ImporteEnBoleta
                                lnDebe = 0
                                If lbGeneraReciboPago = True Then
                                   lnSaldo = -rsReporte.Fields!ImporteEnBoleta
                                Else
                                   lnSaldo = 0
                                End If
                           Else
                                lnDebe = 0
                                If lbGeneraReciboPago = True Then
                                   lnSaldo = rsReporte.Fields!ImporteEnBoleta
                                   lnPago = -rsReporte.Fields!ImporteEnBoleta
                                Else
                                   lnSaldo = 0
                                   lnPago = 0
                                End If
                           End If
                        End If

                        If rsReporte.Fields!idProducto <> lnIdPagosACuenta Then
                           lnTSubTotal = lnTSubTotal + rsReporte.Fields("TotalPagar").Value
                           lnTSubTotalDEBE = lnTSubTotalDEBE + lnDebe
                        End If
                        lnTSubTotalEXO = lnTSubTotalEXO + lnEXO
                        lnTSubTotalSIS = lnTSubTotalSIS + lnSIS
                        lnTSubTotalSOAT = lnTSubTotalSOAT + lnSOAT
                        lnTsubTotalConvenio = lnTsubTotalConvenio + lnConvenio
                        lnTSubTotalPAGO = lnTSubTotalPAGO + lnPago
                        lnTSubTotalSALDO = lnTSubTotalSALDO + lnSaldo
                        
                        If rsReporte.Fields!idProducto <> lnIdPagosACuenta Then
                           lnTotal = lnTotal + rsReporte.Fields("TotalPagar").Value
                           lnTotalDEBE = lnTotalDEBE + lnDebe
                        End If
                        lnTotalSIS = lnTotalSIS + lnSIS
                        lnTotalSOAT = lnTotalSOAT + lnSOAT
                        lnTotalEXO = lnTotalEXO + lnEXO
                        lnTotalConvenio = lnTotalConvenio + lnConvenio
                        lnTotalPAGO = lnTotalPAGO + lnPago
                        lnTotalSALDO = lnTotalSALDO + lnSaldo
                        
                        If rsReporte.Fields!idProducto = lnIdPagosACuenta Then   'Pagos a cuenta
                           lnTotalCredito = lnTotalCredito + rsReporte.Fields("TotalPorPagar").Value
                        End If
                        
                    End If
                    rsReporte.MoveNext
                    If rsReporte.EOF Then
                       Exit Do
                    End If
                Loop
                If lbEsOpenOffice = True Then
                    Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(5) & CStr(iFila) & ":" & mo_ReglasReportes.BuscaNombreColumna(iCol + 8) & CStr(iFila))
                    mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                    If FuenteFinanciamiento = 1 Or FuenteFinanciamiento = 5 Then
                        Call Feuille.getcellbyposition(4, iFila - 1).setFormula(Format(lnTSubTotal, "0.00"))
                    Else
                        Call Feuille.getcellbyposition(4, iFila - 1).setFormula(Format(0, "0.00"))
                    End If
                    Call Feuille.getcellbyposition(5, iFila - 1).setFormula(Format(lnTSubTotalEXO, "0.00"))
                    Call Feuille.getcellbyposition(6, iFila - 1).setFormula(Format(lnTSubTotalSIS, "0.00"))
                    Call Feuille.getcellbyposition(7, iFila - 1).setFormula(Format(lnTSubTotalSOAT, "0.00"))
                    Call Feuille.getcellbyposition(8, iFila - 1).setFormula(Format(lnTsubTotalConvenio, "0.00"))
                    Call Feuille.getcellbyposition(9, iFila - 1).setFormula(Format(lnTSubTotalDEBE, "0.00"))
                Else
                    mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 5, iFila, iCol + 8
                    If FuenteFinanciamiento = 1 Or FuenteFinanciamiento = 5 Then
                      oWorkSheet.Cells(iFila, 5).Value = Format(lnTSubTotal, "0.00")
                    Else
                      oWorkSheet.Cells(iFila, 5).Value = Format(0, "0.00")
                    End If
                    oWorkSheet.Cells(iFila, 6).Value = Format(lnTSubTotalEXO, "0.00")
                    oWorkSheet.Cells(iFila, 7).Value = Format(lnTSubTotalSIS, "0.00")
                    oWorkSheet.Cells(iFila, 8).Value = Format(lnTSubTotalSOAT, "0.00")
                    oWorkSheet.Cells(iFila, 9).Value = Format(lnTsubTotalConvenio, "0.00")
                    oWorkSheet.Cells(iFila, 10).Value = Format(lnTSubTotalDEBE, "0.00")
                End If
                iFila = iFila + 1
             Loop
        End If
        iFila = iFila + 1
        If lbEsOpenOffice = True Then
            Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(2) & CStr(iFila) & ":" & mo_ReglasReportes.BuscaNombreColumna(iCol + 8) & CStr(iFila))
            mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
            Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Total: ")
            If FuenteFinanciamiento = 1 Or FuenteFinanciamiento = 5 Then
                Call Feuille.getcellbyposition(4, iFila - 1).setFormula(Format(lnTotal, "0.00"))
            Else
                Call Feuille.getcellbyposition(4, iFila - 1).setFormula(Format(0, "0.00"))
            End If
            Call Feuille.getcellbyposition(5, iFila - 1).setFormula(Format(lnTotalEXO, "0.00"))
            Call Feuille.getcellbyposition(6, iFila - 1).setFormula(Format(lnTotalSIS, "0.00"))
            Call Feuille.getcellbyposition(7, iFila - 1).setFormula(Format(lnTotalSOAT, "0.00"))
            Call Feuille.getcellbyposition(8, iFila - 1).setFormula(Format(lnTotalConvenio, "0.00"))
            Call Feuille.getcellbyposition(9, iFila - 1).setFormula(Format(lnTotalDEBE, "0.00"))
        Else
            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, iCol + 8
            oWorkSheet.Cells(iFila, 2).Value = "Total: "
            If FuenteFinanciamiento = 1 Or FuenteFinanciamiento = 5 Then
              oWorkSheet.Cells(iFila, 5).Value = Format(lnTotal, "0.00")
            Else
              oWorkSheet.Cells(iFila, 5).Value = Format(0, "0.00")
            End If
            oWorkSheet.Cells(iFila, 6).Value = Format(lnTotalEXO, "0.00")
            oWorkSheet.Cells(iFila, 7).Value = Format(lnTotalSIS, "0.00")
            oWorkSheet.Cells(iFila, 8).Value = Format(lnTotalSOAT, "0.00")
            oWorkSheet.Cells(iFila, 9).Value = Format(lnTotalConvenio, "0.00")
            oWorkSheet.Cells(iFila, 10).Value = Format(lnTotalDEBE, "0.00")
        End If
        
        If lbGeneraReciboPago = True Then
            If lnTotalPagosAdelantados > lnPagoEnFarmacia Then
               lnTotalPagosAdelantados = lnTotalPagosAdelantados - lnPagoEnFarmacia
               lnPagoEnFarmacia = 0
            Else
               lnPagoEnFarmacia = lnPagoEnFarmacia - lnTotalPagosAdelantados
               lnTotalPagosAdelantados = 0
            End If
            lnPagoEnServicios = lnPagoEnServicios - lnTotalPagosAdelantados
        Else
           Select Case mo_ReglasFacturacion.TiposFinanciamientoDevuelveComoSeTrabajaEnEstadoCuenta(lnIdTipoFinanciamientoActual, oConexion)
           Case sghTrabajaSeguroSIS
               lnTotalSIS = lnTotalSIS - lnTotalPagosAdelantados
           Case sghTrabajaSeguroSOAT
               lnTotalSOAT = lnTotalSOAT - lnTotalPagosAdelantados
           Case sghTrabajaSeguroConvenios
               lnTotalConvenio = lnTotalConvenio - lnTotalPagosAdelantados
           End Select
        End If
        lnPagoEnServicios = lnTotalSALDO - lnPagoEnFarmacia
        
        iFila = iFila + 2
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(1, iFila - 1).setFormula("TOTAL PAGO A CUENTA: ")
            Call Feuille.getcellbyposition(3, iFila - 1).setFormula(Format(lnTotalPagosAdelantados, "0.00"))
            iFila = iFila + 1
            Call Feuille.getcellbyposition(1, iFila - 1).setFormula("TOTAL PACIENTE: ")
            Call Feuille.getcellbyposition(3, iFila - 1).setFormula(Format(lnTotalDEBE, "0.00"))
        Else
            oWorkSheet.Cells(iFila, 2).Value = "TOTAL PAGO A CUENTA: "
            oWorkSheet.Cells(iFila, 4).Value = Format(lnTotalPagosAdelantados, "0.00")
            iFila = iFila + 1
            oWorkSheet.Cells(iFila, 2).Value = "TOTAL PACIENTE: "
            oWorkSheet.Cells(iFila, 4).Value = Format(lnTotalDEBE, "0.00") '
        End If
        iFila = iFila + 1
        
        If lbEsOpenOffice = True Then
            If lbGeneraReciboPago = True Then
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("PAGADO: ")
                Call Feuille.getcellbyposition(3, iFila - 1).setFormula(Format(lnTotalPAGO - lnTotalPagosAdelantados + lnPagosXdevoluciones, "0.00"))
                iFila = iFila + 1
                iFila = iFila + 1
                iFila = iFila + 1
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("POR PAGAR PACIENTE: ")
                Call Feuille.getcellbyposition(3, iFila - 1).setFormula(Format(Val(txtTotalApagar.Text), "0.00"))
            Else
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("PAGADO: ")
                Call Feuille.getcellbyposition(3, iFila - 1).setFormula(Format(lnTotalPAGO - lnTotalPagosAdelantados, "0.00"))
                iFila = iFila + 1
                iFila = iFila + 1
                iFila = iFila + 1
                lnTotalSALDO = Val(txtTotalServicios.Text) + Val(txtTotalFarmacia.Text)
                If lnTotalSALDO > 0 Then
                    Call Feuille.getcellbyposition(1, iFila - 1).setFormula("POR PAGAR PACIENTE: ")
                    Call Feuille.getcellbyposition(3, iFila - 1).setFormula(Format(lnTotalSALDO, "0.00"))
                    iFila = iFila + 1
                End If
            End If
        Else
            If lbGeneraReciboPago = True Then
                oWorkSheet.Cells(iFila, 2).Value = "PAGADO: "
                oWorkSheet.Cells(iFila, 4).Value = Format(lnTotalPAGO - lnTotalPagosAdelantados + lnPagosXdevoluciones, "0.00") 'oWorkSheet.Cells(iFila, 4).Value = Format((lnTotalDEBE - Adelanto - lnTotalPAGO), "0.00")
                iFila = iFila + 1
                iFila = iFila + 1
                iFila = iFila + 1
                oWorkSheet.Cells(iFila, 2).Value = "POR PAGAR PACIENTE: "
                oWorkSheet.Cells(iFila, 4).Value = Format(Val(txtTotalApagar.Text), "0.00")  'oWorkSheet.Cells(iFila, 4).Value = Format(lnTotalSALDO + lnTotalPagosAdelantados, "0.00")
            Else
                oWorkSheet.Cells(iFila, 2).Value = "PAGADO: "
                oWorkSheet.Cells(iFila, 4).Value = Format(lnTotalPAGO - lnTotalPagosAdelantados, "0.00") 'oWorkSheet.Cells(iFila, 4).Value = Format((lnTotalDEBE - Adelanto - lnTotalPAGO), "0.00")
                iFila = iFila + 1
                iFila = iFila + 1
                iFila = iFila + 1
                lnTotalSALDO = Val(txtTotalServicios.Text) + Val(txtTotalFarmacia.Text)
                If lnTotalSALDO > 0 Then
                    oWorkSheet.Cells(iFila, 2).Value = "POR PAGAR PACIENTE: "
                    oWorkSheet.Cells(iFila, 4).Value = Format(lnTotalSALDO, "0.00") 'oWorkSheet.Cells(iFila, 4).Value = Format(lnTotalSALDO + lnTotalPagosAdelantados, "0.00")
                    iFila = iFila + 1
                End If
            End If
        End If
        '
        If lbEsOpenOffice = True Then
            If lnTotalSOAT > 0 Then
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("TOTAL A PAGAR SOAT: ")
                Call Feuille.getcellbyposition(3, iFila - 1).setFormula(Format(lnTotalSOAT, "0.00"))
            ElseIf lnTotalSIS > 0 Then
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("TOTAL A PAGAR SIS: ")
                Call Feuille.getcellbyposition(3, iFila - 1).setFormula(Format(lnTotalSIS, "0.00"))
            ElseIf lnTotalConvenio > 0 Then
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("TOTAL A PAGAR CONVENIO: ")
                Call Feuille.getcellbyposition(3, iFila - 1).setFormula(Format(lnTotalConvenio, "0.00"))
            End If
        Else
            If lnTotalSOAT > 0 Then
              oWorkSheet.Cells(iFila, 2).Value = "TOTAL A PAGAR SOAT: "
              oWorkSheet.Cells(iFila, 4).Value = Format(lnTotalSOAT, "0.00")
            ElseIf lnTotalSIS > 0 Then
              oWorkSheet.Cells(iFila, 2).Value = "TOTAL A PAGAR SIS: "
              oWorkSheet.Cells(iFila, 4).Value = Format(lnTotalSIS, "0.00")
            ElseIf lnTotalConvenio > 0 Then
              oWorkSheet.Cells(iFila, 2).Value = "TOTAL A PAGAR CONVENIO: "
              oWorkSheet.Cells(iFila, 4).Value = Format(lnTotalConvenio, "0.00")
            End If
        End If
        '
        If lcListaDeOrdenesDePago <> "" Then
            iFila = iFila + 2
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("* El CAJERO debe emitir Boletas usando " & lcListaDeOrdenesDePago)
            Else
                oWorkSheet.Cells(iFila, 2).Value = "* El CAJERO debe emitir Boletas usando " & lcListaDeOrdenesDePago
            End If
        End If
        '***Donaciones en Farmacia
        If rsItemsDonaciones.RecordCount > 0 Then
           Dim lnCantidadDona As Long, lcCodigoDona As String, lbContinua As Boolean
           iFila = iFila + 2
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("LISTA DE DONACIONES:")
                iFila = iFila + 1
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Descripción")
                Call Feuille.getcellbyposition(8, iFila - 1).setFormula("Cantidad")
                Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(2) & CStr(iFila) & ":" & mo_ReglasReportes.BuscaNombreColumna(9) & CStr(iFila))
                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
            Else
                oWorkSheet.Cells(iFila, 2).Value = "LISTA DE DONACIONES:"
                iFila = iFila + 1
                oWorkSheet.Cells(iFila, 2).Value = "Descripción"
                oWorkSheet.Cells(iFila, 9).Value = "Cantidad"
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 9
            End If
           Set rsReporte = Nothing
           With rsReporte
                  .Fields.Append "Codigo", adVarChar, 10, adFldIsNullable
                  .Fields.Append "Nombre", adVarChar, 200, adFldIsNullable
                  .Fields.Append "Cantidad", adInteger
                  .CursorType = adOpenDynamic
                  .LockType = adLockOptimistic
                  .Open
           End With
           rsItemsDonaciones.MoveFirst
           Do While Not rsItemsDonaciones.EOF
              lbContinua = True
              If rsReporte.RecordCount > 0 Then
                 rsReporte.MoveFirst
                 rsReporte.Find "Codigo='" & rsItemsDonaciones.Fields!Codigo & "'"
                 If Not rsReporte.EOF Then
                    lbContinua = False
                 End If
              End If
              If lbContinua = True Then
                  rsReporte.AddNew
                  rsReporte.Fields!Codigo = rsItemsDonaciones.Fields!Codigo
                  rsReporte.Fields!nombre = rsItemsDonaciones.Fields!nombre
              End If
              rsReporte.Fields!cantidad = rsReporte.Fields!cantidad + rsItemsDonaciones.Fields!cantidad
              rsReporte.Update
              rsItemsDonaciones.MoveNext
           Loop
           rsReporte.Sort = "nombre,codigo"
           rsReporte.MoveFirst
           Do While Not rsReporte.EOF
              iFila = iFila + 1
              lcCodigoDona = rsReporte.Fields!Codigo
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(1, iFila - 1).setFormula(Trim(rsReporte.Fields!Codigo) & " " & rsReporte.Fields!nombre)
                    Call Feuille.getcellbyposition(8, iFila - 1).setFormula(rsReporte.Fields!cantidad)
                Else
                    oWorkSheet.Cells(iFila, 2).Value = Trim(rsReporte.Fields!Codigo) & " " & rsReporte.Fields!nombre
                    oWorkSheet.Cells(iFila, 9).Value = rsReporte.Fields!cantidad
                End If
              rsReporte.MoveNext
           Loop
        End If
        '
        If lbEsOpenOffice = True Then
            Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
            PrintArea(0).Sheet = 0
            PrintArea(0).startcolumn = 0
            PrintArea(0).StartRow = 0
            PrintArea(0).EndColumn = 10
            PrintArea(0).EndRow = iFila
            Call Feuille.SetPrintAreas(PrintArea())
            Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
            MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
        Else
            oWorkSheet.PageSetup.PrintTitleRows = "$1:$7"
            If oWorkSheet.PageSetup.PrintArea <> "" Then oWorkSheet.PageSetup.PrintArea = "$A$1:$K$" & (iFila + 2) 'sighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
            oExcel.Visible = True
            oWorkSheet.PrintPreview
        End If
            
    End If
    MousePointer = 1
    'debb-jamo
    If txtFegreso.Text = "" And ml_IdTipoServicio = sghHospitalizacion Then
       LimparDatos
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set rsReporte = Nothing
    
    If lbEsOpenOffice = True Then
        'Liberar Memoria
        Set Plage = Nothing
        Set Feuille = Nothing
        Set Document = Nothing
        Set Desktop = Nothing
        Set ServiceManager = Nothing
        Set Style = Nothing
        Set Border = Nothing
        'encabezado de pagina
        Set PageStyles = Nothing
        Set Sheet = Nothing
        Set StyleFamilies = Nothing
        Set DefPage = Nothing
        Set Htext = Nothing
        Set Hcontent = Nothing
    Else
        'liberar memoria
        Set oExcel = Nothing
        Set oWorkBookPlantilla = Nothing
        Set oWorkBook = Nothing
        Set oWorkSheet = Nothing
    End If

    '
    mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar ml_idCuentaAtencion, False, 0

End Sub







Private Sub grdCuentasPorTipoServicio_DblClick()
    On Error Resume Next
    If optPreVentaServ.Value = True Then
        txtNroOrdenPagoS.Text = oRsCuentasPorTipoServicio.Fields!nroOrdenPago
        txtNroOrdenPagoS_LostFocus
    ElseIf optExoneracionesFarmacia.Value = True Then
        txtDctoExoneracionFarm = oRsCuentasPorTipoServicio.Fields!NroDocumento
        txtDctoExoneracionFarm_LostFocus
    Else
        txtNroCuenta.Text = oRsCuentasPorTipoServicio.Fields!idCuentaAtencion
        txtNroCuenta_LostFocus
    End If
    TabBusqueda.Tab = 0
    grdCuentasPorTipoServicio.Visible = False
    MousePointer = 1
End Sub

Private Sub grdCuentasPorTipoServicio_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    On Error Resume Next
    grdCuentasPorTipoServicio.Bands(0).Columns("Plan").Header.Caption = "Fuente Financiamiento/IAFA"
    grdCuentasPorTipoServicio.Bands(0).Columns("T_Financiamiento").Header.Caption = "Producto/Plan"
End Sub



Private Sub grdItemsDonaciones_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    grdItemsDonaciones.Bands(0).Columns("Fecha Atencion").Width = 1200
    grdItemsDonaciones.Bands(0).Columns("Fecha Atencion").Format = sighentidades.DevuelveFechaSoloFormato_DMY_HM
    grdItemsDonaciones.Bands(0).Columns("Punto Carga").Width = 1500
    grdItemsDonaciones.Bands(0).Columns("N° Dcto").Width = 600
    grdItemsDonaciones.Bands(0).Columns("Codigo").Width = 600
    grdItemsDonaciones.Bands(0).Columns("Nombre").Width = 3000
    grdItemsDonaciones.Bands(0).Columns("Cantidad").Format = "###0"
    grdItemsDonaciones.Bands(0).Columns("Cantidad").Width = 600
    grdItemsDonaciones.Bands(0).Columns("precio").Format = "#0.0000"
    grdItemsDonaciones.Bands(0).Columns("precio").Width = 800
    grdItemsDonaciones.Bands(0).Columns("Total").Format = "#0.00"
    grdItemsDonaciones.Bands(0).Columns("Total").Width = 900
    grdItemsDonaciones.Bands(0).Columns("FechaVencimiento").Width = 1200
End Sub

Private Sub grdReembolsoF_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
     Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
     grdReembolsoF.Bands(0).Columns("Ident").Width = 1000
     grdReembolsoF.Bands(0).Columns("Año").Width = 500
     grdReembolsoF.Bands(0).Columns("Mes").Width = 400
     grdReembolsoF.Bands(0).Columns("NroSerie").Width = 1000
     grdReembolsoF.Bands(0).Columns("NroDocumento").Width = 1400
     grdReembolsoF.Bands(0).Columns("TipoReembolso").Width = 1200
     grdReembolsoF.Bands(0).Columns("Consumo").Width = 1500
     grdReembolsoF.Bands(0).Columns("Consumo").Format = "###0.00"
     grdReembolsoF.Bands(0).Columns("PorReembolsar").Width = 1500
     grdReembolsoF.Bands(0).Columns("PorReembolsar").Format = "###0.00"
     grdReembolsoF.Bands(0).Columns("Reemb_Servicio").Width = 1500
     grdReembolsoF.Bands(0).Columns("Reemb_Servicio").Format = "###0.00"
     grdReembolsoF.Bands(0).Columns("Reemb_Farmacia").Width = 1500
     grdReembolsoF.Bands(0).Columns("Reemb_Farmacia").Format = "###0.00"
End Sub

Private Sub TabBusqueda_Click(PreviousTab As Integer)
    If PreviousTab = 0 Then
       grdCuentasPorTipoServicio.Visible = True
       UserControl.cmdLiquidacion.Visible = False
       
    Else
       grdCuentasPorTipoServicio.Visible = False
       UserControl.cmdLiquidacion.Visible = True
    End If
End Sub






Private Sub txtDctoExoneracionFarm_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txtDctoExoneracionFarm_LostFocus
    End If
End Sub

Private Sub txtDctoExoneracionFarm_LostFocus()
    If txtDctoExoneracionFarm.Text <> "" Then
        If ml_idUsuarioConPermisoEnSISoEXOoSOAT = 9 Then
            MousePointer = 11
            Dim oRsBuscaDcto As New Recordset
            Dim lcNroMovimiento As String
            Dim lcNroDcto As String
            lcNroDcto = txtDctoExoneracionFarm.Text
            LimparDatos
            txtDctoExoneracionFarm.Text = lcNroDcto
            Set oRsBuscaDcto = mo_ReglasFarmacia.farmMovimientoVentasSeleccionarPorNroDocumentoTipoFinanciamiento(txtDctoExoneracionFarm.Text, 9)
            If oRsBuscaDcto.RecordCount = 0 Then
               oRsBuscaDcto.Close
               MsgBox "Esa DOCUMENTO EXONERADO, No existe", vbInformation, "Error"
               MousePointer = 1
               Exit Sub
            End If
            '
            lnIdTipoFinanciamientoActual = sghTipoFinanciamiento.sghPacienteNormal    'contado
            ml_idCuentaAtencion = 0
            btnAceptar.Visible = False
            txtNroHistoria.Text = "Paciente Externo"
            txtPaciente.Text = "Farmacia: " & oRsBuscaDcto.Fields!farmacia
            'LEER DATOS DE BIENES E INSUMOS
            ucFacturacionBienes.LimpiarGrilla
            ucFacturacionBienes.EstadosFacturacion = ""
            ucFacturacionBienes.idTipoFinanciamiento = lnIdTipoFinanciamientoActual
            ucFacturacionBienes.TipoProducto = sghbien
            ucFacturacionBienes.AgruparPor = Val(cmbAgrupar.ItemData(cmbAgrupar.ListIndex))
            ucFacturacionBienes.movNumero = oRsBuscaDcto.Fields!NroMovimiento
            ucFacturacionBienes.CargaProductosPorMovNumero
            ucFacturacionBienes.ActualizaPreciosImportesEnTodosItemsParaSisSoat (ml_idUsuarioConPermisoEnSISoEXOoSOAT)
            txtTotalFarmacia.Text = ucFacturacionBienes.TotalizaPagoDelPaciente
            txtTotalSeguroFarmacia.Text = ucFacturacionBienes.TotalizaPagoDeSeguros
            oRsBuscaDcto.Close
            ucFacturacionBienesInsumos.Tab = 1
            MousePointer = 1
        Else
            MsgBox "Solo Servicio Social podrá usar esta opcion", vbInformation, "Error"
            LimparDatos
        End If
    End If
End Sub

Private Sub txtEstadoCuenta_Change()
  mo_Formulario.HabilitarDeshabilitar txtEstadoCuenta, False
  mo_Formulario.HabilitarDeshabilitar txtEstadoCuenta, True
  If Left(txtEstadoCuenta.Text, 8) = "(Pagado)" Then
    txtEstadoCuenta.BackColor = vbBlue  '&HC0FFFF
    txtEstadoCuenta.ForeColor = vbWhite
  Else
    txtEstadoCuenta.BackColor = vbRed  '&H80C0FF
    txtEstadoCuenta.ForeColor = vbWhite
  End If
End Sub



Private Sub txtFechaFin_LostFocus()
    If Not esfecha(txtFechaFin.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        txtFechaFin.Text = sighentidades.FECHA_VACIA_DMY
        Exit Sub
    End If

End Sub

Private Sub txtFechaInicio_LostFocus()
    If Not esfecha(txtFechaInicio.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        txtFechaInicio.Text = sighentidades.FECHA_VACIA_DMY
        Exit Sub
    End If
End Sub

Private Sub txtFiltroApellPat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtFiltroApellPat.Text <> "" Then
       oRsCuentasPorTipoServicio.Filter = "ApellidoPaterno like '" & txtFiltroApellPat & "%'"
    End If
End Sub

Private Sub txtNroCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       lcHoraInicioProceso = lcBuscaParametro.RetornaHoraServidorSQL1
       txtNroCuenta_LostFocus
    End If
End Sub

'debb-20/10/2015
Private Sub txtNroCuenta_LostFocus()
10        On Error GoTo ErrFocus
20        If mo_Teclado.TextoEsSoloNumeros(txtNroCuenta.Text) Then
30          MousePointer = 11
            '
40          If oConexionConsulta.State = 1 Then
50             oConexionConsulta.Close
60          End If
70          oConexionConsulta.Open sighentidades.CadenaConexion
80          oConexionConsulta.CursorLocation = adUseClient
90          oConexionConsulta.CommandTimeout = 150
            '
            Dim rs As New Recordset
100         Set rs = mo_AdminAdmision.AtencionesFiltraDatosCabecera(Val(txtNroCuenta.Text), oConexionConsulta)
110         LimparDatos
            
120         If rs.RecordCount > 0 Then
130            lcdTipoFinanciamiento = Trim(rs!dTipoFinanciamiento)
140            ml_idPaciente = rs!idPaciente
150            ml_idTipoSexo = rs!idTipoSexo
160            txtPaciente.Text = rs!ApellidoPaterno + " " + rs!ApellidoMaterno + " " + rs!PrimerNombre + " " + _
                                  IIf(IsNull(rs!SegundoNombre), "", " " + rs!SegundoNombre)
170            txtNroHistoria.Text = rs!NroHistoriaClinica
180            mo_cmbIdTipoGenHistoriaClinica.BoundText = rs!IdTipoNumeracion
               '
190            btnImprimir.Enabled = True
200            cmdImprimeCtaPorServicioHosp.Enabled = True
210            bntLiquidacion.Enabled = True
               '
220            ml_idCuentaAtencion = rs!idCuentaAtencion
230            ml_lbEsPacienteExterno = rs!EsPacienteExterno
240            txtEstadoCuenta.Text = rs!estadoCta & IIf(ml_lbEsPacienteExterno = True, " <> PAC_EXTERNO", "")
250            ml_idAtencion = rs.Fields!idAtencion
260            ml_IdTipoServicio = rs.Fields!idTipoServicio
270            ml_idEstadoCuentaAtencion = rs.Fields!IdEstado
280            If Not IsNull(rs.Fields!IdCondicionAlta) Then
290               ml_dCondicionAlta = Trim(Str(rs.Fields!IdCondicionAlta))
300            End If
310            EncontroCuenta
320            lcHoraFinalProceso = lcBuscaParametro.RetornaHoraServidorSQL1
330            On Error Resume Next
340            lblTiempoDeCargaDeCuenta.Caption = DateDiff("s", CDate(lcHoraInicioProceso), CDate(lcHoraFinalProceso))
              
350         End If
            'oConexionConsulta.Close
360         Set rs = Nothing
370         MousePointer = 1
380      End If
390      Exit Sub
ErrFocus:
          MsgBox Err.Number & " " & Err.Description & _
          sighentidades.DevuelveFuenteDeLineaDelError(Erl(), "Private Sub txtNroCuenta_LostFocus", "ucEstadoCuenta.ctl")   'debb-02/05/2016

End Sub

Private Sub txtNroHistoria_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       btnLeerProductos_Click
    End If
End Sub





Private Sub txtNroOrdenPagoS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txtNroOrdenPagoS_LostFocus
    End If
End Sub

Private Sub txtNroOrdenPagoS_LostFocus()
     If txtNroOrdenPagoS.Text <> "" Then
        btnAceptar.Visible = False
        If ml_idUsuarioConPermisoEnSISoEXOoSOAT = 0 Or ml_idUsuarioConPermisoEnSISoEXOoSOAT = 9 Then
            MousePointer = 11
            Dim lnIdOrdenS As Long
            Dim oRsBuscaCta As New Recordset
            Dim oDoFactOrdenServPagos As New DoFactOrdenServPagos
            lnIdOrdenS = Val(txtNroOrdenPagoS.Text)
            LimparDatos
            txtNroOrdenPagoS.Text = lnIdOrdenS
            oDoFactOrdenServPagos.IdOrden = 0
            Set oDoFactOrdenServPagos = mo_ReglasFacturacion.FactOrdenServicioPagosSeleccionarPorIdOrdenPago(Val(txtNroOrdenPagoS.Text))
            If oDoFactOrdenServPagos.IdOrden > 0 Then
               lnIdOrdenS = oDoFactOrdenServPagos.IdOrden
            Else
               MousePointer = 1
               Exit Sub
            End If
            Set oDoFactOrdenServPagos = Nothing
            'busca si tiene IdCuentaAtencion
            Set oRsBuscaCta = mo_AdminCaja.FactOrdenServicioSeleccionarPuntoCargaPorIdOrden(lnIdOrdenS)
            If oRsBuscaCta.RecordCount > 0 Then
               If oRsBuscaCta.Fields!idCuentaAtencion > 0 Then
                  MsgBox "Esa ORDEN DE PAGO, tiene N° Cuenta", vbInformation, "Error"
                  MousePointer = 1
                  Exit Sub
               End If
            End If
            oRsBuscaCta.Close
            Set oRsBuscaCta = Nothing
            '
            lnIdTipoFinanciamientoActual = sghTipoFinanciamiento.sghPacienteNormal    'contado
            ml_idCuentaAtencion = 0
            If ml_idUsuarioConPermisoEnSISoEXOoSOAT = 9 Then
                btnAceptar.Visible = True
                btnAceptar.Caption = "Actualiza Exoneraciones"
                btnAceptar.Enabled = True
                bntLiquidacion.Visible = True
            End If
            'LEER DATOS DE SERVICIOS
            ucFacturacionServicios.LimpiarGrilla
            ucFacturacionServicios.EstadosFacturacion = ""
            ucFacturacionServicios.idTipoFinanciamiento = lnIdTipoFinanciamientoActual
            ucFacturacionServicios.TipoProducto = sghServicio
            ucFacturacionServicios.idCuentaAtencion = ml_idCuentaAtencion
            ucFacturacionServicios.AgruparPor = Val(cmbAgrupar.ItemData(cmbAgrupar.ListIndex))
            ucFacturacionServicios.IdOrden = lnIdOrdenS
            ucFacturacionServicios.CargaProductosPorIdOrden
            ucFacturacionServicios.ActualizaPreciosImportesEnTodosItemsParaSisSoat (ml_idUsuarioConPermisoEnSISoEXOoSOAT)
            txtTotalServicios.Text = ucFacturacionServicios.TotalizaPagoDelPaciente
            txtTotalSeguroServicio.Text = ucFacturacionServicios.TotalizaPagoDeSeguros
            txtPaciente.Text = "(N° Ord.Pago: " & Trim(txtNroOrdenPagoS.Text) & ")   (N° Orden: " & Trim(Str(lnIdOrdenS)) & ")"
            txtNroHistoria.Text = "Paciente Externo"
            ucFacturacionBienesInsumos.Tab = 0
            MousePointer = 1
        Else
            MsgBox "Solo Servicio Social podrá usar esta opcion", vbInformation, "Error"
            LimparDatos
        End If
     Else
        
     End If

End Sub

Private Sub ucFacturacionBienes_Totalizado(TotalIngresado As Double, TotalPendientePago As Double, TotalPagoACuenta As Double, TotalExonerado As Double, dTotalPagado As Double, dTotalPorDevolver As Double, dTotalDevuelto As Double, dTotalAnulado As Double)
        
        txtIngresadoBien = IIf(TotalIngresado = 0, "", Format(TotalIngresado, "0.00"))
        txtPendientePagoBien = IIf(TotalPendientePago = 0, "", Format(TotalPendientePago, "0.00"))
        txtExoneradoBien = IIf(TotalExonerado = 0, "", Format(TotalExonerado, "0.00"))
        md_TotalBien = TotalIngresado + TotalPendientePago - TotalPagoACuenta - TotalExonerado
        txtTotalBien.Text = IIf(md_TotalBien = 0, "", Format(md_TotalBien, "0.00"))
        
        md_TotalBienPagado = dTotalPagado
        txtTotalBienPagado.Text = IIf(md_TotalBienPagado = 0, "", Format(md_TotalBienPagado, "0.00"))
        
        md_Total = md_TotalBien + md_TotalServ
        txtTotal.Text = IIf(md_Total = 0, "", Format(md_Total, "0.00"))
        
End Sub




Private Sub ucFacturacionServicios_Totalizado(TotalIngresado As Double, TotalPendientePago As Double, TotalPagoACuenta As Double, TotalExonerado As Double, dTotalPagado As Double, dTotalPorDevolver As Double, dTotalDevuelto As Double, dTotalAnulado As Double)
        
        txtIngresadoServ = IIf(TotalIngresado = 0, "", Format(TotalIngresado, "0.00"))
        txtPendientePagoServ = IIf(TotalPendientePago = 0, "", Format(TotalPendientePago, "0.00"))
        txtPagoACuentaServ.Text = IIf(TotalPagoACuenta = 0, "", Format(TotalPagoACuenta, "0.00"))
        txtExoneradoServ = IIf(TotalExonerado = 0, "", Format(TotalExonerado, "0.00"))
        TxtDctosServicio.Text = DevuelveTotalDctosPorIdCuentaAtencion(Val(txtCuenta.Text))
        
        md_TotalServ = TotalIngresado + TotalPendientePago - TotalPagoACuenta - TotalExonerado - TxtDctosServicio
        
        txtTotalServ.Text = IIf(md_TotalServ = 0, "", Format(md_TotalServ, "0.00"))
        txtTotalServPagado.Text = IIf(dTotalPagado = 0, "", Format(dTotalPagado, "0.00"))

        md_Total = md_TotalBien + md_TotalServ
        md_TotalPagado = dTotalPagado + md_TotalBienPagado
        
        txtTotal.Text = IIf(md_Total = 0, "", Format(md_Total, "0.00"))
        txtTotalPagado.Text = IIf(md_TotalPagado = 0, "", Format(md_TotalPagado, "0.00"))

End Sub










Sub CargaTextos()
    txtCtaPagada.Text = "Es usado cuando el Paciente ya Pagó sus Boletas (Farmacia y/o Servicio). El Sistema pone FECHA ALTA ADMINISTRATIVA y desocupa CAMA."
    txtCtaAnulada.Text = "Es usado cuando el Paciente FUGO, pasado el tiempo Justifico en Economía. O tambien cuando se apertura una Cuenta por ERROR. El Sistema pone FECHA ALTA ADMINISTRATIVA y desocupa CAMA."
    txtCtaCerrar.Text = "Es usado cuando el Paciente FUGA, sin cancelar. Se usará como ALERTA cuando el Paciente regrese. El Sistema pone FECHA ALTA ADMINISTRATIVA, deuda Pendiente y desocupa CAMA."
    txtCtaAbrir.Text = "Es usado cuando se quiere agregar un Consumo más o para que el Paciente FUGADO pueda cancelar su deuda. Elimina las FECHA DE ALTA ADMINISTRATIVA."
    txtRecalculo.Text = "Es usado cuando se cambia el 'Plan de Atención' a un Paciente."
    txtPendienteSeguro.Text = "Es usado cuando el Paciente tiene algun SEGURO (Sis, Soat, etc). Cuando aún el Seguro no reembolsa lo gastado por el Paciente. El Sistema pone FECHA ALTA ADMINISTRATIVA, deuda del SEGURO y desocupa CAMA. "
    txtCtaConGarante.Text = "Es usado cuando un GARANTE se compromete a PAGAR la deuda del Paciente. El Sistema pone FECHA ALTA ADMINISTRATIVA y desocupa CAMA."
End Sub



Sub CargaCuentaEnResumen()
    Dim RsBusqueda As New Recordset
    Dim oRsTmp As New Recordset
    Dim oRsTmp1 As New Recordset
    Dim lbNuevo As Boolean
    Dim lcLlave As String
    Dim lcTexto As String
    Dim lnTotalApagar As Double
    Dim lnImpo As Double: Dim lnPrec As Double: Dim lnCant As Long
    Dim lnidReceta As Long
    lnEstadoFacturacionAtendidoOpreventa = sghAtendido
    'crea temporal
    CreaTemporales
    'servicios
    lnTotalApagar = 0
    Set RsBusqueda = ucFacturacionServicios.FacturacionProductos.Clone()
    RsBusqueda.Filter = "idEstadoFacturacion = 1 or idEstadoFacturacion = 4 or idEstadoFacturacion=" & sghConPreVenta
    If RsBusqueda.RecordCount > 0 Then
       RsBusqueda.MoveFirst
       Do While Not RsBusqueda.EOF
          If RsBusqueda.Fields!idestadofacturacion = sghConPreVenta Then
             lnEstadoFacturacionAtendidoOpreventa = sghConPreVenta
          End If
          Set oRsTmp = mo_ReglasComunes.FactPuntosCargaSeleccionarPorFiltro("idPuntoCarga=" & RsBusqueda.Fields!idPuntoCarga, oConexionConsulta)
          lcTexto = ""
          If oRsTmp.RecordCount > 0 Then
             lcTexto = Trim(oRsTmp.Fields!Descripcion)
          End If
          oRsTmp.Close
          '
          lnidReceta = 0
          Set oRsTmp1 = mo_ReglasComunes.RecetaCabeceraXidCuentaAtencion(ml_idCuentaAtencion, oConexionConsulta)
          '
          Select Case mo_ReglasFacturacion.FuentesFinanciamientoDevuelveIdTipoConceptoFarmacia(lnIdPlanActual, oConexionConsulta)
          Case 13     'SIS
               lnCant = RsBusqueda.Fields!CantidadSIS
               lnPrec = RsBusqueda.Fields!precioSIS
               lnImpo = RsBusqueda.Fields!ImporteSIS
          Case 14     'SOAT
               lnCant = RsBusqueda.Fields!CantidadSOAT
               lnPrec = RsBusqueda.Fields!PrecioSOAT
               lnImpo = RsBusqueda.Fields!ImporteSOAT
          Case 23     'Convenio FOSPOLIS
               lnCant = RsBusqueda.Fields!cantidadConv
               lnPrec = RsBusqueda.Fields!precioCONV
               lnImpo = RsBusqueda.Fields!ImporteConv
          Case Else
               lnCant = RsBusqueda.Fields!cantidad
               lnPrec = RsBusqueda.Fields!PrecioUnitario
               lnImpo = RsBusqueda.Fields!TotalPorPagar
          End Select
          '
          lcLlave = lcTexto & " - " & RsBusqueda.Fields!FechaOrden & " - " & RsBusqueda.Fields!nroDcto
          lbNuevo = True
          If oRsCuentaCabecera.RecordCount > 0 Then
             oRsCuentaCabecera.MoveFirst
             oRsCuentaCabecera.Find "llave='" & lcLlave & "'"
             If Not oRsCuentaCabecera.EOF Then
                lbNuevo = False
             End If
          End If
          If lbNuevo Then
                oRsCuentaCabecera.AddNew
                oRsCuentaCabecera.Fields!llave = lcLlave
                oRsCuentaCabecera.Fields!puntoDeCarga = lcTexto
                oRsCuentaCabecera.Fields!fecha = RsBusqueda.Fields!FechaDespacho
                oRsCuentaCabecera.Fields!Servicio = RsBusqueda.Fields!ServicioDeEstancia
                oRsCuentaCabecera.Fields!Importe = lnImpo
                oRsCuentaCabecera.Fields!NroDocumento = RsBusqueda.Fields!nroDcto      'RsBusqueda.Fields!IdOrden
                oRsTmp1.Filter = "idPuntoCarga=" & RsBusqueda.Fields!idPuntoCarga & " and DocumentoDespacho='" & RsBusqueda.Fields!nroDcto & "'"
                If oRsTmp1.RecordCount > 0 Then
                   oRsCuentaCabecera.Fields!Receta = Trim(Str(oRsTmp1.Fields!idReceta))
                End If
          Else
                oRsCuentaCabecera.Fields!Importe = oRsCuentaCabecera.Fields!Importe + lnImpo
          End If
          oRsCuentaCabecera.Update
          oRsCuentaDetalle.AddNew
          oRsCuentaDetalle.Fields!llave = lcLlave
          oRsCuentaDetalle.Fields!Codigo = RsBusqueda.Fields!Codigo
          oRsCuentaDetalle.Fields!Descripcion = Left(RsBusqueda.Fields!NombreProducto, 50)
          oRsCuentaDetalle.Fields!cantidad = lnCant
          oRsCuentaDetalle.Fields!Precio = lnPrec
          oRsCuentaDetalle.Fields!Importe = lnImpo
          If oRsTmp1.RecordCount > 0 Then
             oRsTmp1.MoveFirst
             oRsTmp1.Find "idItem=" & RsBusqueda.Fields!idProducto
             If Not oRsTmp1.EOF Then
                oRsCuentaDetalle.Fields!cantidadReceta = oRsTmp1.Fields!CantidadPedida
             End If
          End If
          If RsBusqueda.Fields!idestadofacturacion = 4 Then
             oRsCuentaDetalle.Fields!NroDocumento = RsBusqueda.Fields!NroComprobante
          End If
          oRsCuentaDetalle.Update
          lnTotalApagar = lnTotalApagar + lnImpo
          RsBusqueda.MoveNext
       Loop
    End If
    Set RsBusqueda = Nothing
    'farmacia
    Set RsBusqueda = ucFacturacionBienes.FacturacionProductos.Clone()
    RsBusqueda.Filter = "idEstadoFacturacion = 1 or idEstadoFacturacion = 4"
    If RsBusqueda.RecordCount > 0 Then
       RsBusqueda.MoveFirst
       Do While Not RsBusqueda.EOF
          Set oRsTmp = mo_ReglasComunes.FactPuntosCargaSeleccionarPorFiltro(" idPuntoCarga=" & RsBusqueda.Fields!idPuntoCarga, oConexionConsulta)
          lcTexto = ""
          If oRsTmp.RecordCount > 0 Then
             lcTexto = Trim(oRsTmp.Fields!Descripcion)
          End If
          oRsTmp.Close
          '
          lnidReceta = 0
          Set oRsTmp1 = mo_ReglasComunes.RecetaCabeceraXidCuentaAtencion(ml_idCuentaAtencion, oConexionConsulta)
          '
          Select Case mo_ReglasFacturacion.FuentesFinanciamientoDevuelveIdTipoConceptoFarmacia(lnIdPlanActual, oConexionConsulta)
          Case 13     'SIS
               lnCant = RsBusqueda.Fields!CantidadSIS
               lnPrec = RsBusqueda.Fields!precioSIS
               lnImpo = RsBusqueda.Fields!ImporteSIS
          Case 14     'SOAT
               lnCant = RsBusqueda.Fields!CantidadSOAT
               lnPrec = RsBusqueda.Fields!PrecioSOAT
               lnImpo = RsBusqueda.Fields!ImporteSOAT
          Case 23     'Convenio FOSPOLIS
               lnCant = RsBusqueda.Fields!cantidadConv
               lnPrec = RsBusqueda.Fields!precioCONV
               lnImpo = RsBusqueda.Fields!ImporteConv
          Case Else
               lnCant = RsBusqueda.Fields!cantidad
               lnPrec = RsBusqueda.Fields!PrecioUnitario
               lnImpo = RsBusqueda.Fields!TotalPorPagar
          End Select
          '
          lcLlave = lcTexto & " - " & RsBusqueda.Fields!FechaOrden
          lbNuevo = True
          If oRsCuentaCabecera.RecordCount > 0 Then
             oRsCuentaCabecera.MoveFirst
             oRsCuentaCabecera.Find "llave='" & lcLlave & "'"
             If Not oRsCuentaCabecera.EOF Then
                lbNuevo = False
             End If
          End If
          If lbNuevo Then
                oRsCuentaCabecera.AddNew
                oRsCuentaCabecera.Fields!llave = lcLlave
                oRsCuentaCabecera.Fields!puntoDeCarga = lcTexto
                oRsCuentaCabecera.Fields!fecha = RsBusqueda.Fields!FechaDespacho
                oRsCuentaCabecera.Fields!Servicio = RsBusqueda.Fields!ServicioDeEstancia
                oRsCuentaCabecera.Fields!Importe = lnImpo
                oRsCuentaCabecera.Fields!NroDocumento = RsBusqueda.Fields!NroDocumento
                oRsTmp1.Filter = "idPuntoCarga=" & RsBusqueda.Fields!idPuntoCarga & " and DocumentoDespacho='" & Trim(RsBusqueda.Fields!nroDcto) & "'"
                If oRsTmp1.RecordCount > 0 Then
                   oRsCuentaCabecera.Fields!Receta = Trim(Str(oRsTmp1.Fields!idReceta))
                End If
          Else
                oRsCuentaCabecera.Fields!Importe = oRsCuentaCabecera.Fields!Importe + lnImpo
          End If
          oRsCuentaCabecera.Update
          oRsCuentaDetalle.AddNew
          oRsCuentaDetalle.Fields!llave = lcLlave
          oRsCuentaDetalle.Fields!Codigo = RsBusqueda.Fields!Codigo
          oRsCuentaDetalle.Fields!Descripcion = Left(RsBusqueda.Fields!NombreProducto, 150)
          oRsCuentaDetalle.Fields!cantidad = lnCant
          oRsCuentaDetalle.Fields!Precio = lnPrec
          oRsCuentaDetalle.Fields!Importe = lnImpo
          oRsCuentaDetalle.Fields!CantDevuelta = RsBusqueda.Fields!CantidadDevuelta
          If oRsTmp1.RecordCount > 0 Then
             oRsTmp1.MoveFirst
             oRsTmp1.Find "idItem=" & RsBusqueda.Fields!idProducto
             If Not oRsTmp1.EOF Then
                oRsCuentaDetalle.Fields!cantidadReceta = oRsTmp1.Fields!CantidadPedida
             End If
          End If
          If RsBusqueda.Fields!idestadofacturacion = 4 Then
             oRsCuentaDetalle.Fields!NroDocumento = RsBusqueda.Fields!NroComprobante
          End If
          oRsCuentaDetalle.Update
          lnTotalApagar = lnTotalApagar + lnImpo
          RsBusqueda.MoveNext
       Loop
    End If
    Set RsBusqueda = Nothing
    txtTotalApagar.Text = lnTotalApagar - Val(txtPagosAdelantoC.Text)
    oRsCuentaCabecera.Sort = "fecha"
    '
    Set grdCabecera.DataSource = oRsCuentaCabecera
    Set grdDetalle.DataSource = Nothing
    Set oRsTmp1 = Nothing
    Set RsBusqueda = Nothing
    Set oRsTmp = Nothing
End Sub

Sub CreaTemporales()
    On Error GoTo ErrTmp
    Dim lnLinea As Integer
    lnLinea = 1
    If oRsCuentaCabecera.State = 1 Then Set oRsCuentaCabecera = Nothing
    With oRsCuentaCabecera
        .Fields.Append "llave", adVarChar, 150, adFldIsNullable    'PuntoDeCarga+Fecha+NroDocumento
        .Fields.Append "PuntoDeCarga", adVarChar, 50, adFldIsNullable
        .Fields.Append "Fecha", adDate, , adFldIsNullable
        .Fields.Append "NroDocumento", adVarChar, 20, adFldIsNullable
        .Fields.Append "Servicio", adVarChar, 60, adFldIsNullable
        .Fields.Append "Importe", adDouble
        .Fields.Append "Receta", adVarChar, 10, adFldIsNullable
        .LockType = adLockOptimistic
        .Open
    End With
    lnLinea = 2
    If oRsCuentaDetalle.State = 1 Then Set oRsCuentaDetalle = Nothing
    With oRsCuentaDetalle
        .Fields.Append "llave", adVarChar, 70, adFldIsNullable    'PuntoDeCarga+Fecha
        .Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
        .Fields.Append "Descripcion", adVarChar, 250, adFldIsNullable
        .Fields.Append "CantDevuelta", adDouble
        .Fields.Append "Cantidad", adDouble
        .Fields.Append "Precio", adDouble
        .Fields.Append "Importe", adDouble
        .Fields.Append "NroDocumento", adVarChar, 30, adFldIsNullable
        .Fields.Append "CantidadReceta", adDouble
        .LockType = adLockOptimistic
        .Open
    End With
    Exit Sub
ErrTmp:
    If Err.Number = 3219 Then
       If lnLinea = 1 Then
          oRsCuentaCabecera.Close
       Else
          oRsCuentaDetalle.Close
       End If
       Resume
    Else
       MsgBox Err.Description
    End If
End Sub

Private Sub grdCabecera_DblClick()
    grdDetalle.Caption = "Punto de Carga: " & oRsCuentaCabecera.Fields!llave
    oRsCuentaDetalle.Filter = "llave='" & Trim(oRsCuentaCabecera.Fields!llave) & "'"
    Set grdDetalle.DataSource = oRsCuentaDetalle
End Sub
Private Sub grdCabecera_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = 13 Then
       grdCabecera_DblClick
    End If
End Sub

Private Sub grdDetalle_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
     Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
     grdDetalle.Bands(0).Columns("llave").Hidden = True
     grdDetalle.Bands(0).Columns("Codigo").Width = 1000
     grdDetalle.Bands(0).Columns("Descripcion").Width = 4000
     grdDetalle.Bands(0).Columns("Cantidad").Width = 1000
     grdDetalle.Bands(0).Columns("Cantidad").Format = "###0"
     grdDetalle.Bands(0).Columns("Precio").Width = 1500
     grdDetalle.Bands(0).Columns("Precio").Format = "#0.0000"
     grdDetalle.Bands(0).Columns("Importe").Width = 1500
     grdDetalle.Bands(0).Columns("Importe").Format = "#0.00"
End Sub
Private Sub grdCabecera_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
     Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
     grdCabecera.Bands(0).Columns("llave").Hidden = True
     grdCabecera.Bands(0).Columns("puntoDeCarga").Width = 2500
     grdCabecera.Bands(0).Columns("Fecha").Width = 1500
     grdCabecera.Bands(0).Columns("Fecha").Format = sighentidades.DevuelveFechaSoloFormato_DMY_HM
     grdCabecera.Bands(0).Columns("Servicio").Width = 4500
     grdCabecera.Bands(0).Columns("Importe").Width = 1500
     grdCabecera.Bands(0).Columns("Importe").Format = "#0.00"
End Sub



Public Sub ConsultaDetalleCuenta(lnIdCuentaAtencion As Long)
    txtNroCuenta.Text = Trim(Str(lnIdCuentaAtencion))
    txtNroCuenta_LostFocus
    TabBusqueda.Tab = 0
    ucFacturacionBienesInsumos.Tab = 2
    grdCuentasPorTipoServicio.Visible = False
End Sub




Private Sub cmdExoneracion_Click()
'  Dim oExcel As Excel.Application
'  Dim oDOPaciente As New doPaciente
'    Dim oWorkBookPlantilla As Workbook
'    Dim oWorkBook As Workbook
'    Dim oWorkSheet As Worksheet
'    Dim iFila As Long: Dim iCol As Integer
'    Dim rsReporte As New Recordset
'    Dim ms_EstadosFacturacion As String
'    Dim ms_TiposFinanciamiento As String
'    Dim ml_AgruparPor As Long
'    Dim mo_ReporteUtil As New ReporteUtil
'    Dim idPuntoCarga As Long: Dim lnIdTipoServicio As Long
'    Dim lcEstancia As String
'
'    Dim lnTSubTotal As Double: Dim lnTSubTotalSIS As Double: Dim lnTSubTotalEXO As Double
'    Dim lnTSubTotalPAGO As Double: Dim lnTSubTotalDEBE As Double: Dim lnTSubTotalSALDO As Double
'
'    Dim lnTotal As Double: Dim lnTotalSIS As Double: Dim lnTotalEXO As Double
'    Dim lnTotalPAGO As Double: Dim lnTotalDEBE As Double: Dim lnTotalSALDO As Double
'
'    Dim lnDebe As Double: Dim lnPago As Double: Dim lnSaldo As Double
'    Dim lnTDebe As Double: Dim lnTPago As Double: Dim lnTSaldo As Double
'    Dim lnSIS As Double: Dim lnEXO As Double: Dim lnTotalCredito  As Double: Dim lnSOAT As Double
'    Dim lcBuscaParametro As New SIGHDatos.Parametros
'    Dim CantidadSOAT As Long: Dim PrecioSOAT As Double
'    If txtPaciente.Text = "" Then
'        MsgBox "Tiene que LEER el Estado de Cuenta del Paciente", vbInformation, "Resultado"
'    Else
'        MousePointer = 11
'        'Crea nueva hoja
'        Set oExcel = GalenhosExcelApplication()  'New Excel.Application
'        Set oWorkBook = oExcel.Workbooks.Add
'        'Abre, copia y cierra la plantilla
'        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\EExoneracion.xls")
'        oWorkBookPlantilla.Worksheets("Exoneracion").Copy Before:=oWorkBook.Sheets(1)
'        oWorkBookPlantilla.Close
'        'Activa la primera hoja
'        Set oWorkSheet = oWorkBook.Sheets(1)
'        oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\Imagen de reportes.jpg"
'        '*******************************************Inicio de Reporte
'        Set rsReporte = mo_ReglasFacturacion.AtencionesFiltraEgresosPorIdAtencion(ml_idAtencion)
'        If Val(txtNroHistoria.Text) > 0 Then
'           Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorHistoriaClinicaDefinitiva(CLng(Trim(txtNroHistoria.Text)))
'        End If
'        If rsReporte.RecordCount > 0 Then
'        oWorkSheet.Cells(3, 2).Value = "Paciente: " & txtPaciente.Text
'        oWorkSheet.Cells(3, 6).Value = "Nro Historia: " & Trim(txtNroHistoria.Text)
'        oWorkSheet.Cells(4, 2).Value = "Dirección: " & oDOPaciente.DireccionDomicilio
'        oWorkSheet.Cells(4, 6).Value = "Nº Cuenta: " & txtCuenta.Text
'        oWorkSheet.Cells(5, 2).Value = "Ocupación: " & mo_reglasComunes.DescripcionOcupacion(oDOPaciente.idTipoOcupacion)
'        oWorkSheet.Cells(5, 6).Value = "Fecha: " & Format(Now, SIGHENTIDADES.DevuelveFechaSoloFormato_DMY_HMS)
'        oWorkSheet.Cells(11, 2).Value = "SERVICIO DE EGRESO: " & txtServicio.Text
'
'        End If
'        rsReporte.Close
'
'        iFila = 15
'        iCol = 2
'        Select Case ml_idUsuarioConPermisoEnSISoEXOoSOAT
'        Case 9   'EXONERACIONES
'            oWorkSheet.Cells(iFila - 2, 5).Value = "MONTO EXONERADO (S/.)"
'        End Select
'        lnTotal = 0: lnTotalSIS = 0: lnTotalEXO = 0: lnTotalPAGO = 0: lnTotalDEBE = 0: lnTotalSALDO = 0: lnTotalCredito = 0
'        lnTDebe = 0: lnTPago = 0: lnTSaldo = 0
'
'
'        oWorkSheet.Cells(iFila, 3).Value = "MONTO LIQUIDACIÓN BRUTO"
'        iFila = iFila + 1
'        oWorkSheet.Cells(iFila - 1, 6).Value = Format(txtTotalApagar.Text, "0.00")
'
'        iFila = iFila + 1
'        oWorkSheet.Cells(iFila, 3).Value = "DESCUÉNTESE EN LOS RUBROS:"
'        iFila = iFila + 1
'        'Farmacia
'        Set rsReporte = ucFacturacionBienes.FacturacionProductos
'        If rsReporte.RecordCount > 0 Then
'            rsReporte.MoveFirst
'            oWorkSheet.Cells(iFila, 2).Value = iFila - 17
'            oWorkSheet.Cells(iFila, 3).Value = "Farmacia"
'            iFila = iFila + 1
'            lnTSubTotal = 0: lnTSubTotalSIS = 0: lnTSubTotalEXO = 0
'            lnTSubTotalPAGO = 0: lnTSubTotalDEBE = 0: lnTSubTotalSALDO = 0
'            Do While Not rsReporte.EOF
'                    Select Case ml_idUsuarioConPermisoEnSISoEXOoSOAT
'                    Case 9   'EXONERACIONES
'                        lnSOAT = IIf(IsNull(rsReporte.Fields!importeEXO), 0, rsReporte.Fields!importeEXO)
'                        CantidadSOAT = 0
'                        PrecioSOAT = 0
'                    End Select
'                    lnTSubTotal = lnTSubTotal + lnSOAT
'                    lnTotal = lnTotal + lnSOAT
'
'                rsReporte.MoveNext
'             Loop
'             oWorkSheet.Cells(iFila - 1, 5).Value = Format(lnTSubTotal, "0.00")
'             oWorkSheet.Cells(iFila - 1, 7).Value = SIGHENTIDADES.Numlet(SIGHENTIDADES.DevuelveNumeroSinDecimales(lnTSubTotal)) & " CON " & SIGHENTIDADES.DevuelveSoloDecimales(lnTSubTotal) & "/100"
'        End If
'        rsReporte.Close
'
'        'Servicios
'        Set rsReporte = ucFacturacionServicios.FacturacionProductos
'        If rsReporte.RecordCount > 0 Then
'            rsReporte.Sort = "IdPuntoCarga"
'            rsReporte.MoveFirst
'            Do While Not rsReporte.EOF
'                oWorkSheet.Cells(iFila, 2).Value = iFila - 17
'                idPuntoCarga = rsReporte.Fields("IdPuntoCarga").Value
'                oWorkSheet.Cells(iFila, 3).Value = mo_reglasComunes.FactPuntosCargaSeleccionarPorIdDevDescripcion(rsReporte.Fields("IdPuntoCarga").Value)
'                iFila = iFila + 1
'                lnTSubTotal = 0: lnTSubTotalSIS = 0: lnTSubTotalEXO = 0
'                lnTSubTotalPAGO = 0: lnTSubTotalDEBE = 0: lnTSubTotalSALDO = 0
'                Do While Not rsReporte.EOF And idPuntoCarga = rsReporte.Fields("IdPuntoCarga").Value
'                        Select Case ml_idUsuarioConPermisoEnSISoEXOoSOAT
'                        Case 9   'EXONERACIONES
'                            lnSOAT = IIf(IsNull(rsReporte.Fields!importeEXO), 0, rsReporte.Fields!importeEXO)
'                            CantidadSOAT = 0
'                            PrecioSOAT = 0
'                        End Select
'                        lnTSubTotal = lnTSubTotal + lnSOAT
'                        lnTotal = lnTotal + lnSOAT
'
'                    rsReporte.MoveNext
'                    If rsReporte.EOF Then Exit Do
'                Loop
''                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 3, iFila, 8
'                oWorkSheet.Cells(iFila - 1, 5).Value = Format(lnTSubTotal, "0.00")
'                oWorkSheet.Cells(iFila - 1, 7).Value = SIGHENTIDADES.Numlet(SIGHENTIDADES.DevuelveNumeroSinDecimales(lnTSubTotal)) & " CON " & SIGHENTIDADES.DevuelveSoloDecimales(lnTSubTotal) & "/100"
'             Loop
'        End If
'        iFila = iFila + 1
'        oWorkSheet.Cells(iFila, 3).Value = "TOTAL A EXONERAR (DESCUENTO): "
'        oWorkSheet.Cells(iFila, 6).Value = Format(lnTotal, "0.00")
'        oWorkSheet.Cells(iFila, 7).Value = SIGHENTIDADES.Numlet(SIGHENTIDADES.DevuelveNumeroSinDecimales(lnTotal)) & " CON " & SIGHENTIDADES.DevuelveSoloDecimales(lnTotal) & "/100"
'
'        iFila = iFila + 1
'        iFila = iFila + 1
'        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 9
'        oWorkSheet.Cells(iFila, 3).Value = "TOTAL A PAGAR: "
'        oWorkSheet.Cells(iFila, 6).Value = Format((Val(oWorkSheet.Cells(15, 6).Value) - (Val(oWorkSheet.Cells(iFila - 1, 6).Value))), "0.00")
'        oWorkSheet.Cells(iFila, 7).Value = SIGHENTIDADES.Numlet(SIGHENTIDADES.DevuelveNumeroSinDecimales(Val(oWorkSheet.Cells(iFila, 6).Value))) & " CON " & SIGHENTIDADES.DevuelveSoloDecimales(Val(oWorkSheet.Cells(iFila, 6).Value)) & "/100"
'
'        iFila = iFila + 1
'        iFila = iFila + 1
'        iFila = iFila + 1
'        iFila = iFila + 1
'        iFila = iFila + 1
'        iFila = iFila + 1
'        iFila = iFila + 1
'        iFila = iFila + 1
'        iFila = iFila + 1
'        '
'        If lcListaDeOrdenesDePago <> "" Then
'            iFila = iFila - 1
'            oWorkSheet.Cells(iFila, 3).Value = "* El CAJERO debe emitir Boletas usando " & lcListaDeOrdenesDePago
'        End If
'        '
'        oWorkSheet.PageSetup.PrintTitleRows = "$1:$13"
'        If oWorkSheet.PageSetup.PrintArea <> "" Then oWorkSheet.PageSetup.PrintArea = "$A$1:$J$" & (iFila + 2) 'sighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
'        'oExcel.Visible = True
'        'oWorkSheet.PrintPreview
'        oWorkSheet.PrintOut
'    End If
'    Set oWorkSheet = Nothing
'    Set oExcel = Nothing
'    MousePointer = 1
    Dim oDOPaciente As New doPaciente
    Dim iFila As Long: Dim iCol As Integer
    Dim rsReporte As New Recordset
    Dim ms_EstadosFacturacion As String
    Dim ms_TiposFinanciamiento As String
    Dim ml_AgruparPor As Long
    Dim mo_ReporteUtil As New sighentidades.ReporteUtil
    Dim idPuntoCarga As Long: Dim lnIdTipoServicio As Long
    Dim lcEstancia As String
    
    Dim lnTSubTotal As Double: Dim lnTSubTotalSIS As Double: Dim lnTSubTotalEXO As Double
    Dim lnTSubTotalPAGO As Double: Dim lnTSubTotalDEBE As Double: Dim lnTSubTotalSALDO As Double
    
    Dim lnTotal As Double: Dim lnTotalSIS As Double: Dim lnTotalEXO As Double
    Dim lnTotalPAGO As Double: Dim lnTotalDEBE As Double: Dim lnTotalSALDO As Double
    
    Dim lnDebe As Double: Dim lnPago As Double: Dim lnSaldo As Double
    Dim lnTDebe As Double: Dim lnTPago As Double: Dim lnTSaldo As Double
    Dim lnSIS As Double: Dim lnEXO As Double: Dim lnTotalCredito  As Double: Dim lnSOAT As Double
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim CantidadSOAT As Long: Dim PrecioSOAT As Double, lnNroExoneracion As Long
    Dim lbEsOpenOffice As Boolean
    Dim lcSql As String
    
    lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
    
    If txtPaciente.Text = "" Then
        MsgBox "Tiene que LEER el Estado de Cuenta del Paciente", vbInformation, "Resultado"
    Else
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
            Dim lnHWnd As Long
        Else
            Dim oExcel As Excel.Application
            Dim oWorkBookPlantilla As Workbook
            Dim oWorkBook As Workbook
            Dim oWorkSheet As Worksheet
            Dim oRange As range
            Dim range As Excel.range
            Dim borders As Excel.borders
        End If
        
        MousePointer = 11
        If lbEsOpenOffice = True Then
            'Abre el archivo ExcelOpenOffice
            lcArchivoExcel = App.Path + "\Plantillas\EExoneracion.ods"
    '        FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
    '        Chemin = "file:///" & App.Path & "\Plantillas\"
    '        Chemin = Replace(Chemin, "\", "/")
    '        Fichier = Chemin & "/OpenOffice.ods"
            '
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
            'Encabezado de Pagina
            mo_CabeceraReportes.CabeceraReportes Document, True
            ' Pone la ventana en primer plano, pasándole el Hwnd
            ret = SetForegroundWindow(lnHWnd)
        Else
            'Crea nueva hoja
            Set oExcel = GalenhosExcelApplication()  'New Excel.Application
            Set oWorkBook = oExcel.Workbooks.Add
            'Abre, copia y cierra la plantilla
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\EExoneracion.xls")
            oWorkBookPlantilla.Worksheets("Exoneracion").Copy Before:=oWorkBook.Sheets(1)
            oWorkBookPlantilla.Close
            'Activa la primera hoja
            Set oWorkSheet = oWorkBook.Sheets(1)
            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
        End If
        '*******************************************Inicio de Reporte
        Set rsReporte = mo_ReglasFacturacion.AtencionesFiltraEgresosPorIdAtencion(ml_idAtencion)
        If Val(txtNroHistoria.Text) > 0 Then
           Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorHistoriaClinicaDefinitiva(CLng(Trim(txtNroHistoria.Text)))
        End If
        
        lnNroExoneracion = mo_ReglasFacturacion.NroExoneracionXcuenta(CLng(txtCuenta.Text))
        
        If rsReporte.RecordCount > 0 Then
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(1, 0).setFormula("EXONERACIÓN Nº " & lnNroExoneracion)
                Call Feuille.getcellbyposition(1, 2).setFormula("Paciente: " & txtPaciente.Text)
                Call Feuille.getcellbyposition(5, 2).setFormula("Nro Historia: " & Trim(txtNroHistoria.Text))
                Call Feuille.getcellbyposition(1, 3).setFormula("Dirección: " & oDOPaciente.DireccionDomicilio)
                Call Feuille.getcellbyposition(5, 3).setFormula("Nº Cuenta: " & txtCuenta.Text)
                Call Feuille.getcellbyposition(1, 4).setFormula("Ocupación: " & mo_ReglasComunes.DescripcionOcupacion(oDOPaciente.idTipoOcupacion))
                Call Feuille.getcellbyposition(5, 4).setFormula("Fecha: " & Format(Now, sighentidades.DevuelveFechaSoloFormato_DMY_HMS))
                Call Feuille.getcellbyposition(1, 10).setFormula("SERVICIO DE EGRESO: " & txtServicio.Text)
            Else
                oWorkSheet.Cells(1, 2).Value = "EXONERACIÓN Nº " & lnNroExoneracion
                oWorkSheet.Cells(3, 2).Value = "Paciente: " & txtPaciente.Text
                oWorkSheet.Cells(3, 6).Value = "Nro Historia: " & Trim(txtNroHistoria.Text)
                oWorkSheet.Cells(4, 2).Value = "Dirección: " & oDOPaciente.DireccionDomicilio
                oWorkSheet.Cells(4, 6).Value = "Nº Cuenta: " & txtCuenta.Text
                oWorkSheet.Cells(5, 2).Value = "Ocupación: " & mo_ReglasComunes.DescripcionOcupacion(oDOPaciente.idTipoOcupacion)
                oWorkSheet.Cells(5, 6).Value = "Fecha: " & Format(Now, sighentidades.DevuelveFechaSoloFormato_DMY_HMS)
                oWorkSheet.Cells(11, 2).Value = "SERVICIO DE EGRESO: " & txtServicio.Text
            End If
        End If
        rsReporte.Close
        
        iFila = 15
        iCol = 2
        Select Case ml_idUsuarioConPermisoEnSISoEXOoSOAT
        Case 9   'EXONERACIONES
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(4, iFila - 3).setFormula("MONTO EXONERADO (S/.)")
            Else
                oWorkSheet.Cells(iFila - 2, 5).Value = "MONTO EXONERADO (S/.)"
            End If
        End Select
        lnTotal = 0: lnTotalSIS = 0: lnTotalEXO = 0: lnTotalPAGO = 0: lnTotalDEBE = 0: lnTotalSALDO = 0: lnTotalCredito = 0
        lnTDebe = 0: lnTPago = 0: lnTSaldo = 0
        
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(2, iFila - 1).setFormula("MONTO LIQUIDACIÓN BRUTO")
            iFila = iFila + 1
            'actualizado 28102014 yamill palomino
            'Call Feuille.getcellbyposition(5, iFila - 2).setFormula(Format(txtTotalApagar.Text, "0.00"))
            Call Feuille.getcellbyposition(5, iFila - 2).setFormula(Format(txtTotalApagar.Text, "0.00"))
            
            iFila = iFila + 1
            Call Feuille.getcellbyposition(2, iFila - 1).setFormula("DESCUÉNTESE EN LOS RUBROS:")
        Else
            oWorkSheet.Cells(iFila, 3).Value = "MONTO LIQUIDACIÓN BRUTO"
            iFila = iFila + 1
            'actualizado 28102014 yamill palomino
            'oWorkSheet.Cells(iFila - 1, 6).Value = Format(txtTotalApagar.Text, "0.00")
            oWorkSheet.Cells(iFila - 1, 6).Value = Format(txtTotalApagar.Text, "0.00")
            
            iFila = iFila + 1
            oWorkSheet.Cells(iFila, 3).Value = "DESCUÉNTESE EN LOS RUBROS:"
        End If
        
        iFila = iFila + 1
        'Farmacia
        Set rsReporte = ucFacturacionBienes.FacturacionProductos
        If rsReporte.RecordCount > 0 Then
            rsReporte.MoveFirst
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula(iFila - 18)
                Call Feuille.getcellbyposition(2, iFila - 1).setFormula("Farmacia")
            Else
                oWorkSheet.Cells(iFila, 2).Value = iFila - 17
                oWorkSheet.Cells(iFila, 3).Value = "Farmacia"
            End If
            
            iFila = iFila + 1
            lnTSubTotal = 0: lnTSubTotalSIS = 0: lnTSubTotalEXO = 0
            lnTSubTotalPAGO = 0: lnTSubTotalDEBE = 0: lnTSubTotalSALDO = 0
            Do While Not rsReporte.EOF
                    Select Case ml_idUsuarioConPermisoEnSISoEXOoSOAT
                    Case 9   'EXONERACIONES
                        lnSOAT = IIf(IsNull(rsReporte.Fields!importeEXO), 0, rsReporte.Fields!importeEXO)
                        CantidadSOAT = 0
                        PrecioSOAT = 0
                    End Select
                    lnTSubTotal = lnTSubTotal + lnSOAT
                    lnTotal = lnTotal + lnSOAT
                    
                rsReporte.MoveNext
             Loop
             If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(4, iFila - 2).setFormula(Format(lnTSubTotal, "0.00"))
                Call Feuille.getcellbyposition(6, iFila - 2).setFormula(sighentidades.Numlet(sighentidades.DevuelveNumeroSinDecimales(lnTSubTotal)) & " CON " & sighentidades.DevuelveSoloDecimales(lnTSubTotal) & "/100")
             Else
                oWorkSheet.Cells(iFila - 1, 5).Value = Format(lnTSubTotal, "0.00")
                oWorkSheet.Cells(iFila - 1, 7).Value = sighentidades.Numlet(sighentidades.DevuelveNumeroSinDecimales(lnTSubTotal)) & " CON " & sighentidades.DevuelveSoloDecimales(lnTSubTotal) & "/100"
             End If
        End If
        rsReporte.Close
        
        'Servicios
        Set rsReporte = ucFacturacionServicios.FacturacionProductos
        If rsReporte.RecordCount > 0 Then
            rsReporte.Sort = "IdPuntoCarga"
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(1, iFila - 1).setFormula(iFila - 18)
                    idPuntoCarga = rsReporte.Fields("IdPuntoCarga").Value
                    Call Feuille.getcellbyposition(2, iFila - 1).setFormula(mo_ReglasComunes.FactPuntosCargaSeleccionarPorIdDevDescripcion(rsReporte.Fields("IdPuntoCarga").Value))
                Else
                    oWorkSheet.Cells(iFila, 2).Value = iFila - 17
                    idPuntoCarga = rsReporte.Fields("IdPuntoCarga").Value
                    oWorkSheet.Cells(iFila, 3).Value = mo_ReglasComunes.FactPuntosCargaSeleccionarPorIdDevDescripcion(rsReporte.Fields("IdPuntoCarga").Value)
                End If
                iFila = iFila + 1
                lnTSubTotal = 0: lnTSubTotalSIS = 0: lnTSubTotalEXO = 0
                lnTSubTotalPAGO = 0: lnTSubTotalDEBE = 0: lnTSubTotalSALDO = 0
                Do While Not rsReporte.EOF And idPuntoCarga = rsReporte.Fields("IdPuntoCarga").Value
                        Select Case ml_idUsuarioConPermisoEnSISoEXOoSOAT
                        Case 9   'EXONERACIONES
                            lnSOAT = IIf(IsNull(rsReporte.Fields!importeEXO), 0, rsReporte.Fields!importeEXO)
                            CantidadSOAT = 0
                            PrecioSOAT = 0
                        End Select
                        lnTSubTotal = lnTSubTotal + lnSOAT
                        lnTotal = lnTotal + lnSOAT
                        
                    rsReporte.MoveNext
                    If rsReporte.EOF Then Exit Do
                Loop
'                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 3, iFila, 8
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(4, iFila - 2).setFormula(Format(lnTSubTotal, "0.00"))
                    Call Feuille.getcellbyposition(6, iFila - 2).setFormula(sighentidades.Numlet(sighentidades.DevuelveNumeroSinDecimales(lnTSubTotal)) & " CON " & sighentidades.DevuelveSoloDecimales(lnTSubTotal) & "/100")
                Else
                    oWorkSheet.Cells(iFila - 1, 5).Value = Format(lnTSubTotal, "0.00")
                    oWorkSheet.Cells(iFila - 1, 7).Value = sighentidades.Numlet(sighentidades.DevuelveNumeroSinDecimales(lnTSubTotal)) & " CON " & sighentidades.DevuelveSoloDecimales(lnTSubTotal) & "/100"
                End If
             Loop
        End If
        iFila = iFila + 1
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(2, iFila - 1).setFormula("TOTAL A EXONERAR (DESCUENTO): ")
            Call Feuille.getcellbyposition(5, iFila - 1).setFormula(Format(lnTotal, "0.00"))
            Call Feuille.getcellbyposition(6, iFila - 1).setFormula(sighentidades.Numlet(sighentidades.DevuelveNumeroSinDecimales(lnTotal)) & " CON " & sighentidades.DevuelveSoloDecimales(lnTotal) & "/100")
        Else
            oWorkSheet.Cells(iFila, 3).Value = "TOTAL A EXONERAR (DESCUENTO): "
            oWorkSheet.Cells(iFila, 6).Value = Format(lnTotal, "0.00")
            oWorkSheet.Cells(iFila, 7).Value = sighentidades.Numlet(sighentidades.DevuelveNumeroSinDecimales(lnTotal)) & " CON " & sighentidades.DevuelveSoloDecimales(lnTotal) & "/100"
        End If
        iFila = iFila + 1
        iFila = iFila + 1
        
'        If lbEsOpenOffice = True Then
'            Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(2) & CStr(iFila) & ":" & mo_ReglasReportes.BuscaNombreColumna(9) & CStr(iFila))
'            mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
'            Call Feuille.getcellbyposition(2, iFila - 1).setFormula("TOTAL A PAGAR: ")
'            Call Feuille.getcellbyposition(5, iFila - 1).setFormula(Format((Val(Feuille.getcellbyposition(5, 14).Value) - (Val(Feuille.getcellbyposition(5, iFila - 1).Value))), "0.00"))
'            Call Feuille.getcellbyposition(6, iFila - 1).setFormula(sighEntidades.Numlet(sighEntidades.DevuelveNumeroSinDecimales(Val(Feuille.getcellbyposition(5, iFila - 1).Value))) & " CON " & sighEntidades.DevuelveSoloDecimales(Val(Feuille.getcellbyposition(5, iFila - 1).Value)) & "/100")
'        Else
'            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 9
'            oWorkSheet.Cells(iFila, 3).Value = "TOTAL A PAGAR: "
'            oWorkSheet.Cells(iFila, 6).Value = Format((Val(oWorkSheet.Cells(15, 6).Value) - (Val(oWorkSheet.Cells(iFila - 2, 6).Value))), "0.00")
'            oWorkSheet.Cells(iFila, 7).Value = sighEntidades.Numlet(sighEntidades.DevuelveNumeroSinDecimales(Val(oWorkSheet.Cells(iFila, 6).Value))) & " CON " & sighEntidades.DevuelveSoloDecimales(Val(oWorkSheet.Cells(iFila, 6).Value)) & "/100"
'        End If
        'debb2014b
        If lbEsOpenOffice = True Then
            Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(2) & CStr(iFila) & ":" & mo_ReglasReportes.BuscaNombreColumna(9) & CStr(iFila))
            mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
            Call Feuille.getcellbyposition(2, iFila - 1).setFormula("TOTAL A PAGAR: ")
            Call Feuille.getcellbyposition(5, iFila - 1).setFormula(Format((Val(Feuille.getcellbyposition(5, 14).Value)), "0.00"))
            Call Feuille.getcellbyposition(6, iFila - 1).setFormula(sighentidades.Numlet(sighentidades.DevuelveNumeroSinDecimales(Val(Feuille.getcellbyposition(5, iFila - 1).Value))) & " CON " & sighentidades.DevuelveSoloDecimales(Val(Feuille.getcellbyposition(5, iFila - 1).Value)) & "/100")
            Call Feuille.getcellbyposition(5, 14).setFormula(Format(Val(Feuille.getcellbyposition(5, 14).Value) + Val(Feuille.getcellbyposition(5, iFila - 3).Value), "0.00"))
        Else
            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 9
            oWorkSheet.Cells(iFila, 3).Value = "TOTAL A PAGAR: "
            oWorkSheet.Cells(iFila, 6).Value = Format(Val(oWorkSheet.Cells(15, 6).Value), "0.00")
            oWorkSheet.Cells(iFila, 7).Value = sighentidades.Numlet(sighentidades.DevuelveNumeroSinDecimales(Val(oWorkSheet.Cells(iFila, 6).Value))) & " CON " & sighentidades.DevuelveSoloDecimales(Val(oWorkSheet.Cells(iFila, 6).Value)) & "/100"
            
            oWorkSheet.Cells(15, 6).Value = Format((Val(oWorkSheet.Cells(15, 6).Value) + (Val(oWorkSheet.Cells(iFila - 2, 6).Value))), "0.00")
        End If
        'debb2014b
        
        iFila = iFila + 1
        iFila = iFila + 1
        iFila = iFila + 1
        iFila = iFila + 1
        iFila = iFila + 1
        iFila = iFila + 1
        iFila = iFila + 1
        iFila = iFila + 1
        iFila = iFila + 1
        '
        If lcListaDeOrdenesDePago <> "" Then
            iFila = iFila - 1
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(2, iFila - 1).setFormula(" El CAJERO debe emitir Boletas usando " & lcListaDeOrdenesDePago)
            Else
                oWorkSheet.Cells(iFila, 3).Value = " El CAJERO debe emitir Boletas usando " & lcListaDeOrdenesDePago
            End If
        End If
        '
        If lbEsOpenOffice = True Then
            Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
            PrintArea(0).Sheet = 0
            PrintArea(0).startcolumn = 1
            PrintArea(0).StartRow = 0
            PrintArea(0).EndColumn = 9
            PrintArea(0).EndRow = iFila
            Call Feuille.SetPrintAreas(PrintArea())
            Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
            MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
        Else
            oWorkSheet.PageSetup.PrintTitleRows = "$1:$13"
            If oWorkSheet.PageSetup.PrintArea <> "" Then oWorkSheet.PageSetup.PrintArea = "$A$1:$J$" & (iFila + 2) 'sighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
            oExcel.Visible = True
            oWorkSheet.PrintPreview
            'oWorkSheet.PrintOut
        End If
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
        'encabezado de pagina
        Set PageStyles = Nothing
        Set Sheet = Nothing
        Set StyleFamilies = Nothing
        Set DefPage = Nothing
        Set Htext = Nothing
        Set Hcontent = Nothing
    Else
        'liberar memoria
        Set oExcel = Nothing
        Set oWorkBookPlantilla = Nothing
        Set oWorkBook = Nothing
        Set oWorkSheet = Nothing
    End If
    MousePointer = 1


End Sub



'******** daniel 18/12/2009 (inicio)
'******** muestra todas las Cuentas de un Paciente en un Combo
'******** se creo comboBox:  cmbCitas
Sub CargaCtasDelPaciente()
          On Error GoTo ErrCargaCtaP        'debb-02/05/2016
          Dim lcSql As String
          Dim oRsTmp As New Recordset
10        Set oRsTmp = mo_AdminAdmision.AtencionesListaCuentasXpaciente(ml_idPaciente, oConexionConsulta)
20        cmbCtas.Clear
30        If oRsTmp.RecordCount > 0 Then
40           oRsTmp.MoveFirst
50           Do While Not oRsTmp.EOF
60              lcSql = Trim(Str(oRsTmp.Fields!idCuentaAtencion)) & " - " & Format(oRsTmp.Fields!FechaIngreso & " " & oRsTmp.Fields!HoraIngreso, sighentidades.DevuelveFechaSoloFormato_DMY_HM) & " - " & oRsTmp.Fields!Descripcion
70              cmbCtas.AddItem lcSql
80              oRsTmp.MoveNext
90           Loop
100       End If
110       oRsTmp.Close
120       Set oRsTmp = Nothing
          Exit Sub      'debb-02/05/2016
ErrCargaCtaP:           'debb-02/05/2016
          MsgBox Err.Number & " " & Err.Description & _
          sighentidades.DevuelveFuenteDeLineaDelError(Erl(), "Sub CargaCtasDelPaciente", "ucEstadoCuenta.ctl")   'debb-02/05/2016
End Sub





Private Sub cmbCtas_Click()
        txtNroCuenta.Text = Trim(Left(cmbCtas.Text, InStr(cmbCtas.Text, "-") - 1))
        lcHoraInicioProceso = lcBuscaParametro.RetornaHoraServidorSQL1
        txtNroCuenta_LostFocus
End Sub
Private Sub cmbCtas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       cmbCtas_Click
    End If
End Sub



Sub CargaGridReembolsos()
          On Error GoTo CargaGridR      'debb-02/05/2016
          Dim lnTreembolsoF As Double, lnTreembolsoS As Double
          Dim lbReembolsoSoloEnFarmacia As Boolean, lbReembolsoSoloEnServicios As Boolean
          Dim lcDocFarmacia As String, lcDocServicio As String
10        Set oRsReembolsos = mo_ReglasFacturacion.FacturacionReembolsosSeleccionarPorCuenta(ml_idCuentaAtencion, oConexionConsulta)
20        Set grdReembolsoF.DataSource = oRsReembolsos
30        lnTreembolsoS = 0: lnTreembolsoF = 0
40        lcDocFarmacia = "": lcDocServicio = ""
50        lbReembolsoSoloEnServicios = False: lbReembolsoSoloEnFarmacia = False
60        If oRsReembolsos.RecordCount > 0 Then
70           oRsReembolsos.MoveFirst
80           Do While Not oRsReembolsos.EOF
90              lnTreembolsoS = lnTreembolsoS + oRsReembolsos.Fields!reemb_servicio
100             lnTreembolsoF = lnTreembolsoF + oRsReembolsos.Fields!reemb_farmacia
110             If oRsReembolsos.Fields!reemb_servicio > 0 Then
120                lbReembolsoSoloEnServicios = True
130             End If
140             If oRsReembolsos.Fields!reemb_farmacia > 0 Then
150                lbReembolsoSoloEnFarmacia = True
160             End If
170             txtPorReembolsar.Text = Format(oRsReembolsos.Fields!porReembolsar, "####,###.#0")
180             If Val(oRsReembolsos.Fields!NroDocumento) > 0 Then
190                If oRsReembolsos.Fields!reemb_servicio > 0 Then
200                   lcDocServicio = oRsReembolsos.Fields!nroSerie + "-" + oRsReembolsos.Fields!NroDocumento
210                End If
220                If oRsReembolsos.Fields!reemb_farmacia > 0 Then
230                   lcDocFarmacia = oRsReembolsos.Fields!nroSerie + "-" + oRsReembolsos.Fields!NroDocumento
240                End If
250             End If
260             oRsReembolsos.MoveNext
270          Loop
280          txtReembolsoF.Text = Format(lnTreembolsoF, "####,###.#0")
290          txtReembolsoS.Text = Format(lnTreembolsoS, "####,###.#0")
300          txtReembolsoT.Text = Format(lnTreembolsoS + lnTreembolsoF, "####,###.#0")
             
310       End If
320       ActualizaEstadoAtencionEnGridServiciosYfarmacia lbReembolsoSoloEnFarmacia, lbReembolsoSoloEnServicios, lcDocFarmacia, lcDocServicio
          Exit Sub      'debb-02/05/2016
CargaGridR:             'debb-02/05/2016
          MsgBox Err.Number & " " & Err.Description & _
          sighentidades.DevuelveFuenteDeLineaDelError(Erl(), "Sub CargaGridReembolsos", "ucEstadoCuenta.ctl")   'debb-02/05/2016
End Sub
'****Debb2104

'Solo cuando el ESTADO DE CUENTA está 'Pagada' o 'Reembolso Parcial'
'****Debb2104
Sub ActualizaEstadoAtencionEnGridServiciosYfarmacia(lbReembolsoSoloEnFarmacia As Boolean, lbReembolsoSoloEnServicios As Boolean, lcDocReembolsoFarmacia As String, lcDocReembolsoServicio As String)
    Select Case ml_idEstadoCuentaAtencion
    Case 15
       If lbReembolsoSoloEnServicios = True Then
          ucFacturacionServicios.ActualizaEstadoAtencionEnGridServiciosYfarmacia ml_idEstadoCuentaAtencion, lcDocReembolsoServicio
       End If
       If lbReembolsoSoloEnFarmacia = True Then
          ucFacturacionBienes.ActualizaEstadoAtencionEnGridServiciosYfarmacia ml_idEstadoCuentaAtencion, lcDocReembolsoFarmacia
       End If
    Case 4
       ucFacturacionServicios.ActualizaEstadoAtencionEnGridServiciosYfarmacia ml_idEstadoCuentaAtencion, lcDocReembolsoServicio
       ucFacturacionBienes.ActualizaEstadoAtencionEnGridServiciosYfarmacia ml_idEstadoCuentaAtencion, lcDocReembolsoFarmacia
    End Select
End Sub


Sub CargaGrillaDonaciones()
          On Error GoTo ErrCargaGrillaD      'debb-02/05/2016
          Dim lcSql As String, lnTotalDona As Double
10        Set rsItemsDonaciones = mo_ReglasFarmacia.DonacionesXcuenta(ml_idCuentaAtencion, oConexionConsulta)
20        lnTotalDona = 0
30        If rsItemsDonaciones.RecordCount > 0 Then
40           rsItemsDonaciones.MoveFirst
50           Do While Not rsItemsDonaciones.EOF
60              lnTotalDona = lnTotalDona + rsItemsDonaciones.Fields!Total
70              rsItemsDonaciones.MoveNext
80           Loop
90           rsItemsDonaciones.MoveFirst
100       End If
110       txtTotalDonaciones.Text = Format(lnTotalDona, "####,###.#0")
120       Set grdItemsDonaciones.DataSource = rsItemsDonaciones
          Exit Sub  'debb-02/05/2016
ErrCargaGrillaD:    'debb-02/05/2016
          MsgBox Err.Number & " " & Err.Description & _
          sighentidades.DevuelveFuenteDeLineaDelError(Erl(), "Sub CargaGrillaDonaciones", "ucEstadoCuenta.ctl")   'debb-02/05/2016
End Sub

Sub InicilizarParametros()
    wxParametro302 = lcBuscaParametro.SeleccionaFilaParametro(302)
End Sub

