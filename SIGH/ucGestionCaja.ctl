VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.UserControl ucGestionCaja 
   ClientHeight    =   9420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13200
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   9420
   ScaleWidth      =   13200
   Begin TabDlg.SSTab tabGestionCaja 
      Height          =   9345
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   13185
      _ExtentX        =   23257
      _ExtentY        =   16484
      _Version        =   393216
      Tab             =   1
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
      TabCaption(0)   =   "Gestión de caja"
      TabPicture(0)   =   "ucGestionCaja.ctx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraCajero"
      Tab(0).Control(1)=   "frmGestionCaja"
      Tab(0).Control(2)=   "frmResumenCaja"
      Tab(0).Control(3)=   "lblNombre"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Registro de comprobante"
      TabPicture(1)   =   "ucGestionCaja.ctx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label8"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "optPagarEstadoDeCTAFarmacia"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "optPagarEstadoDeCuenta"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmbOrdenProvenienteDe"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "OptOrdenFarmacia"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "optRealizarDevolucion"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "optReimprimirComprobante"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "optCobrarOrdenExistente"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtFechaApertura"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtNombres"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "fraPaciente"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Frame3"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Frame4"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Frame5"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "fraOpciones"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "ucFacturacionProductos"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "UcFacturacionContado1"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "frmPreventaServ"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "cmbFechaIngreso"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "cmbIdPuntoDeCarga"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "cmbIdTurno"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "cmbIdCaja"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Frame1"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "cmdPendientesFarmacia"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "cmdPendientesPago"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Frame2"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "tabFactProductosPorCuenta"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "FraServHosp"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "ucGestionCajaFact1"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).ControlCount=   29
      TabCaption(2)   =   "Devolución por Nota de Crédito (F1)"
      TabPicture(2)   =   "ucGestionCaja.ctx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "btnAceptarPagoNota"
      Tab(2).Control(1)=   "fraNotaCredito"
      Tab(2).ControlCount=   2
      Begin SISGalenPlus.ucGestionCajaFact ucGestionCajaFact1 
         Height          =   1545
         Left            =   150
         TabIndex        =   177
         Top             =   5370
         Visible         =   0   'False
         Width           =   12990
         _ExtentX        =   22781
         _ExtentY        =   6906
      End
      Begin VB.Frame fraCajero 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Left            =   -74880
         TabIndex        =   159
         Top             =   840
         Width           =   12900
         Begin VB.CommandButton cmdBuscarPorApell 
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
            Left            =   2940
            Picture         =   "ucGestionCaja.ctx":0054
            Style           =   1  'Graphical
            TabIndex        =   184
            Top             =   510
            Width           =   345
         End
         Begin VB.CommandButton btnLimpiar 
            Height          =   315
            Left            =   11520
            Picture         =   "ucGestionCaja.ctx":05DE
            Style           =   1  'Graphical
            TabIndex        =   183
            Top             =   555
            Width           =   1275
         End
         Begin VB.CommandButton btnBuscar 
            Height          =   315
            Left            =   11505
            Picture         =   "ucGestionCaja.ctx":31BA
            Style           =   1  'Graphical
            TabIndex        =   182
            Top             =   195
            Width           =   1305
         End
         Begin VB.CheckBox chkSoloCredito 
            Alignment       =   1  'Right Justify
            Caption         =   "Solo CREDITO"
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
            Left            =   10050
            TabIndex        =   178
            Top             =   930
            Width           =   1455
         End
         Begin VB.ComboBox cmbIdTurnoBusqueda 
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
            Left            =   5190
            Style           =   2  'Dropdown List
            TabIndex        =   166
            Top             =   510
            Width           =   1125
         End
         Begin VB.ComboBox cmbIdCajaBusqueda 
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
            Left            =   3360
            Style           =   2  'Dropdown List
            TabIndex        =   165
            Top             =   510
            Width           =   1815
         End
         Begin VB.TextBox txtNroSerieBusqueda 
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
            Left            =   120
            MaxLength       =   4
            TabIndex        =   164
            Top             =   510
            Width           =   570
         End
         Begin VB.TextBox txtNroHistoriaBusqueda 
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
            Left            =   1860
            MaxLength       =   10
            TabIndex        =   163
            Top             =   510
            Width           =   1065
         End
         Begin VB.TextBox txtNroDocumentoBusqueda 
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
            Left            =   690
            MaxLength       =   8
            TabIndex        =   162
            Top             =   510
            Width           =   1095
         End
         Begin VB.TextBox TxtRsocial 
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
            Left            =   6330
            MaxLength       =   20
            TabIndex        =   161
            Top             =   510
            Width           =   1335
         End
         Begin VB.ComboBox cmbIdResponsable 
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
            Left            =   690
            Style           =   2  'Dropdown List
            TabIndex        =   160
            Top             =   900
            Width           =   4395
         End
         Begin MSMask.MaskEdBox txtFdesde 
            Height          =   315
            Left            =   7710
            TabIndex        =   167
            Top             =   510
            Width           =   1830
            _ExtentX        =   3228
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
         Begin MSMask.MaskEdBox txtFhasta 
            Height          =   315
            Left            =   9630
            TabIndex        =   168
            Top             =   510
            Width           =   1860
            _ExtentX        =   3281
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   $"ucGestionCaja.ctx":5E03
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
            TabIndex        =   170
            Top             =   210
            Width           =   11715
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Cajero"
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
            TabIndex        =   169
            Top             =   930
            Width           =   510
         End
      End
      Begin VB.Frame frmGestionCaja 
         Caption         =   "Gestión de caja"
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
         Height          =   6135
         Left            =   -74880
         TabIndex        =   156
         Top             =   2250
         Width           =   12975
         Begin UltraGrid.SSUltraGrid grdGestionCaja 
            Height          =   4095
            Left            =   105
            TabIndex        =   157
            Top             =   240
            Width           =   12780
            _ExtentX        =   22543
            _ExtentY        =   7223
            _Version        =   131072
            GridFlags       =   17040384
            LayoutFlags     =   67108884
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Comprobantes de pago"
         End
         Begin UltraGrid.SSUltraGrid grdNotaCredito 
            Height          =   1575
            Left            =   120
            TabIndex        =   158
            Top             =   4440
            Width           =   12780
            _ExtentX        =   22543
            _ExtentY        =   2778
            _Version        =   131072
            GridFlags       =   17040384
            LayoutFlags     =   67108884
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Devolución por Notas de Credito"
         End
      End
      Begin VB.Frame frmResumenCaja 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   -74880
         TabIndex        =   145
         Top             =   8370
         Width           =   12975
         Begin VB.Label lblNroDevNotaCredito 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "..."
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
            Height          =   210
            Left            =   10290
            TabIndex        =   155
            Top             =   510
            Width           =   180
         End
         Begin VB.Label lblTotalDevNotaCredito 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "..."
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
            Height          =   210
            Left            =   10290
            TabIndex        =   154
            Top             =   120
            Width           =   180
         End
         Begin VB.Label lblTotAnulados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   7635
            TabIndex        =   153
            Top             =   120
            Width           =   180
         End
         Begin VB.Label lblTotFacturas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4785
            TabIndex        =   152
            Top             =   120
            Width           =   180
         End
         Begin VB.Label lblTotBoletas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2040
            TabIndex        =   151
            Top             =   120
            Width           =   180
         End
         Begin VB.Label lblNroAnulados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   7635
            TabIndex        =   150
            Top             =   510
            Width           =   180
         End
         Begin VB.Label lblNroDocumentos 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   12675
            TabIndex        =   149
            Top             =   510
            Width           =   180
         End
         Begin VB.Label lblNroFacturas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4785
            TabIndex        =   148
            Top             =   510
            Width           =   180
         End
         Begin VB.Label lblNroBoletas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2040
            TabIndex        =   147
            Top             =   510
            Width           =   180
         End
         Begin VB.Label txtTotalCajero 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   12675
            TabIndex        =   146
            Top             =   120
            Width           =   180
         End
      End
      Begin VB.Frame fraNotaCredito 
         Height          =   7935
         Left            =   -74880
         TabIndex        =   109
         Top             =   360
         Width           =   12975
         Begin VB.Frame Frame8 
            Height          =   1335
            Left            =   120
            TabIndex        =   138
            Top             =   240
            Width           =   12675
            Begin VB.CommandButton btnNotaBusca 
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
               Left            =   3135
               Picture         =   "ucGestionCaja.ctx":5EB5
               Style           =   1  'Graphical
               TabIndex        =   185
               Top             =   825
               Width           =   345
            End
            Begin VB.TextBox txtNotaDocumento 
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
               TabIndex        =   140
               Top             =   840
               Width           =   1815
            End
            Begin VB.TextBox txtNotaSerie 
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
               TabIndex        =   139
               Top             =   840
               Width           =   1215
            End
            Begin Threed.SSOption optCanjeNotaCredito 
               Height          =   255
               Left            =   390
               TabIndex        =   141
               Top             =   240
               Width           =   4305
               _ExtentX        =   7594
               _ExtentY        =   450
               _Version        =   262144
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "DEVOLUCIÓN DE DINERO POR NOTA DE CREDITO"
               Value           =   -1
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               Caption         =   "F1"
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
               Height          =   195
               Left            =   120
               TabIndex        =   144
               Top             =   240
               Width           =   195
            End
            Begin VB.Label Label32 
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
               TabIndex        =   143
               Top             =   600
               Width           =   825
            End
            Begin VB.Label Label33 
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
               TabIndex        =   142
               Top             =   600
               Width           =   1245
            End
         End
         Begin VB.Frame fraNota 
            Caption         =   "Nota de Crédito"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3975
            Left            =   120
            TabIndex        =   110
            Top             =   1680
            Width           =   12735
            Begin VB.CommandButton btnLimpiarNota 
               Caption         =   "Limpiar"
               DisabledPicture =   "ucGestionCaja.ctx":643F
               DownPicture     =   "ucGestionCaja.ctx":6828
               Height          =   585
               Left            =   5040
               Picture         =   "ucGestionCaja.ctx":6C34
               Style           =   1  'Graphical
               TabIndex        =   127
               Top             =   360
               Width           =   1335
            End
            Begin VB.TextBox txtNotaCredAprueba 
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
               Left            =   7350
               MaxLength       =   50
               TabIndex        =   126
               Top             =   1800
               Width           =   5235
            End
            Begin VB.TextBox txtNotaFechaCanje 
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
               Left            =   9960
               TabIndex        =   125
               Top             =   1200
               Width           =   2565
            End
            Begin VB.Frame DetalleNotaCredito 
               BorderStyle     =   0  'None
               Height          =   1215
               Left            =   120
               TabIndex        =   120
               Top             =   2280
               Width           =   12495
               Begin VB.TextBox txtNotaTotal 
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
                  Left            =   10680
                  MultiLine       =   -1  'True
                  TabIndex        =   122
                  Top             =   360
                  Width           =   1815
               End
               Begin VB.TextBox txtNotaConcepto 
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
                  TabIndex        =   121
                  Top             =   360
                  Width           =   10665
               End
               Begin VB.Label Label34 
                  Alignment       =   2  'Center
                  BackColor       =   &H8000000C&
                  Caption         =   "TOTAL(S/.) "
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
                  Left            =   10680
                  TabIndex        =   124
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.Label Label35 
                  Alignment       =   2  'Center
                  BackColor       =   &H8000000C&
                  Caption         =   "CONCEPTO"
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
                  TabIndex        =   123
                  Top             =   0
                  Width           =   10665
               End
            End
            Begin VB.TextBox txtNotaCredDebDireccion 
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
               TabIndex        =   119
               Top             =   1800
               Width           =   4875
            End
            Begin VB.TextBox txtNotaRazonSocial 
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
               TabIndex        =   118
               Top             =   1200
               Width           =   4875
            End
            Begin VB.TextBox txtNotaRuc 
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
               MaxLength       =   50
               TabIndex        =   117
               Top             =   1200
               Width           =   2235
            End
            Begin VB.TextBox txtNotaFechaEmision 
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
               MaxLength       =   50
               TabIndex        =   116
               Top             =   1200
               Width           =   2595
            End
            Begin VB.TextBox txtNotaMotivo 
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
               MaxLength       =   50
               TabIndex        =   115
               Top             =   1800
               Width           =   2295
            End
            Begin VB.TextBox txtSerieNota 
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
               MaxLength       =   3
               TabIndex        =   114
               Top             =   600
               Width           =   1215
            End
            Begin VB.TextBox txtDocumentoNota 
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
               MaxLength       =   12
               TabIndex        =   113
               Top             =   600
               Width           =   1815
            End
            Begin VB.CheckBox chkRevertirPagoNota 
               Caption         =   "DESEA REVERTIR LA DEVOLUCIÓN DE DINERO POR NOTA DE CRÉDITO"
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
               Height          =   375
               Left            =   120
               TabIndex        =   112
               Top             =   3480
               Visible         =   0   'False
               Width           =   6375
            End
            Begin VB.TextBox txtEstadoNota 
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
               MaxLength       =   12
               TabIndex        =   111
               Top             =   600
               Width           =   1880
            End
            Begin VB.Label Label43 
               Caption         =   "Responsable de la Aprobación de la Nota de Crédito"
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
               TabIndex        =   137
               Top             =   1560
               Width           =   4845
            End
            Begin VB.Label Label31 
               Caption         =   "Fecha de Canje"
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
               Left            =   9960
               TabIndex        =   136
               Top             =   960
               Width           =   1605
            End
            Begin VB.Label Label36 
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
               TabIndex        =   135
               Top             =   960
               Width           =   1845
            End
            Begin VB.Label Label37 
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
               TabIndex        =   134
               Top             =   1560
               Width           =   1245
            End
            Begin VB.Label Label38 
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
               TabIndex        =   133
               Top             =   1560
               Width           =   4845
            End
            Begin VB.Label Label39 
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
               TabIndex        =   132
               Top             =   960
               Width           =   1245
            End
            Begin VB.Label Label40 
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
               TabIndex        =   131
               Top             =   960
               Width           =   4845
            End
            Begin VB.Label Label25 
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
               TabIndex        =   130
               Top             =   360
               Width           =   1245
            End
            Begin VB.Label Label41 
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
               TabIndex        =   129
               Top             =   360
               Width           =   825
            End
            Begin VB.Label Label42 
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
               TabIndex        =   128
               Top             =   360
               Width           =   1245
            End
         End
      End
      Begin VB.CommandButton btnAceptarPagoNota 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "ucGestionCaja.ctx":7040
         DownPicture     =   "ucGestionCaja.ctx":74A0
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -74760
         Picture         =   "ucGestionCaja.ctx":7915
         Style           =   1  'Graphical
         TabIndex        =   108
         Top             =   8400
         Width           =   1455
      End
      Begin VB.Frame FraServHosp 
         Height          =   2325
         Left            =   4470
         TabIndex        =   95
         Top             =   360
         Width           =   1875
         Begin Threed.SSOption optSHtodos 
            Height          =   255
            Left            =   450
            TabIndex        =   96
            Top             =   150
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   450
            _Version        =   262144
            BackStyle       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Todos"
            Value           =   -1
         End
         Begin Threed.SSOption optSHLaboratorio 
            Height          =   255
            Left            =   450
            TabIndex        =   97
            Top             =   525
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   262144
            BackStyle       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Laboratorio"
         End
         Begin Threed.SSOption optSHrayosX 
            Height          =   255
            Left            =   450
            TabIndex        =   98
            Top             =   900
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   450
            _Version        =   262144
            BackStyle       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Rayos X"
         End
         Begin Threed.SSOption optSHtomografia 
            Height          =   255
            Left            =   450
            TabIndex        =   99
            Top             =   1275
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   450
            _Version        =   262144
            BackStyle       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Tomografía"
         End
         Begin Threed.SSOption optSHecogGeneral 
            Height          =   255
            Left            =   450
            TabIndex        =   100
            Top             =   1650
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   450
            _Version        =   262144
            BackStyle       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Ecog.General"
         End
         Begin Threed.SSOption optSHecogObst 
            Height          =   255
            Left            =   450
            TabIndex        =   101
            Top             =   2025
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   450
            _Version        =   262144
            BackStyle       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Ecogr.Obst"
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "F8"
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
            Height          =   195
            Left            =   75
            TabIndex        =   105
            Top             =   540
            Width           =   195
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "F9"
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
            Height          =   195
            Left            =   75
            TabIndex        =   104
            Top             =   900
            Width           =   195
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "F11"
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
            Height          =   195
            Left            =   75
            TabIndex        =   103
            Top             =   1665
            Width           =   300
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "F12"
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
            Height          =   195
            Left            =   75
            TabIndex        =   102
            Top             =   2040
            Width           =   300
         End
      End
      Begin TabDlg.SSTab tabFactProductosPorCuenta 
         Height          =   3300
         Left            =   120
         TabIndex        =   39
         Top             =   4200
         Visible         =   0   'False
         Width           =   12810
         _ExtentX        =   22595
         _ExtentY        =   5821
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
         TabCaption(0)   =   "Servicios"
         TabPicture(0)   =   "ucGestionCaja.ctx":7D8A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label23"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label24"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "ucFactServiciosPorCuenta"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "txtCtaServExonerado"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "txtCtaServTservicio"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "Farmacia"
         TabPicture(1)   =   "ucGestionCaja.ctx":7DA6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtCtaFarmTfarmacia"
         Tab(1).Control(1)=   "txtCtaFarmExonerado"
         Tab(1).Control(2)=   "ucFactBienesPorCuenta"
         Tab(1).Control(3)=   "Label27"
         Tab(1).Control(4)=   "Label26"
         Tab(1).ControlCount=   5
         Begin VB.TextBox txtCtaFarmTfarmacia 
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
            Height          =   345
            Left            =   -63660
            TabIndex        =   88
            Top             =   2910
            Width           =   1200
         End
         Begin VB.TextBox txtCtaFarmExonerado 
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
            Height          =   345
            Left            =   -73890
            TabIndex        =   87
            Top             =   2910
            Width           =   1200
         End
         Begin VB.TextBox txtCtaServTservicio 
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
            Height          =   345
            Left            =   11400
            TabIndex        =   86
            Top             =   2910
            Width           =   1200
         End
         Begin VB.TextBox txtCtaServExonerado 
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
            Height          =   345
            Left            =   1170
            TabIndex        =   85
            Top             =   2880
            Width           =   1200
         End
         Begin SISGalenPlus.ucFactItemsPorCuenta2 ucFactBienesPorCuenta 
            Height          =   2460
            Left            =   -74880
            TabIndex        =   41
            Top             =   420
            Width           =   12600
            _ExtentX        =   22225
            _ExtentY        =   4339
         End
         Begin SISGalenPlus.ucFactItemsPorCuenta2 ucFactServiciosPorCuenta 
            Height          =   2520
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   12570
            _ExtentX        =   22172
            _ExtentY        =   4445
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Total Farmacia:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -64980
            TabIndex        =   84
            Top             =   2970
            Width           =   1305
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Exonerado:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74850
            TabIndex        =   83
            Top             =   2970
            Width           =   1065
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Total Servicio:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   10170
            TabIndex        =   82
            Top             =   2970
            Width           =   1200
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Exonerado:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   81
            Top             =   2970
            Width           =   1065
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2325
         Left            =   180
         TabIndex        =   43
         Top             =   360
         Width           =   4275
         Begin VB.CommandButton cmdOrdenExistenteFS 
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
            Left            =   3540
            Picture         =   "ucGestionCaja.ctx":7DC2
            Style           =   1  'Graphical
            TabIndex        =   179
            Top             =   570
            Width           =   345
         End
         Begin Threed.SSOption optNuevoOrdenPagoConHistoria 
            Height          =   255
            Left            =   360
            TabIndex        =   44
            Top             =   150
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   450
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "NUEVO PAGO CON H. CLINICA/N° CUENTA"
         End
         Begin Threed.SSOption optNuevoOrdenPagoSinHistoria 
            Height          =   255
            Left            =   360
            TabIndex        =   45
            Top             =   2010
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   450
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "PAGOS ADMINIST/SERVICIOS INTERMEDIOS"
         End
         Begin Threed.SSOption optRealizarAnulacion 
            Height          =   255
            Left            =   360
            TabIndex        =   46
            Top             =   1080
            Width           =   2325
            _ExtentX        =   4101
            _ExtentY        =   450
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "ANULAR COMPROBANTE"
         End
         Begin Threed.SSOption optOrdenExistenteFS 
            Height          =   255
            Left            =   360
            TabIndex        =   72
            Top             =   615
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   450
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "PAGAR ORDEN EXISTENTE (Serv/Farm)"
         End
         Begin Threed.SSOption optPagarCtaTotal 
            Height          =   255
            Left            =   360
            TabIndex        =   80
            Top             =   1545
            Width           =   3105
            _ExtentX        =   5477
            _ExtentY        =   450
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "PAGAR CUENTA TOTAL (Serv/Farm)"
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "F7"
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
            Height          =   195
            Left            =   105
            TabIndex        =   56
            Top             =   2025
            Width           =   195
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "F6"
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
            Height          =   195
            Left            =   105
            TabIndex        =   55
            Top             =   1530
            Width           =   195
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "F5"
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
            Height          =   195
            Left            =   105
            TabIndex        =   54
            Top             =   1080
            Width           =   195
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "F4"
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
            Height          =   195
            Left            =   105
            TabIndex        =   53
            Top             =   615
            Width           =   195
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "F3"
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
            Height          =   195
            Left            =   105
            TabIndex        =   52
            Top             =   150
            Width           =   195
         End
      End
      Begin VB.CommandButton cmdPendientesPago 
         Caption         =   "..."
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
         Left            =   2850
         TabIndex        =   78
         ToolTipText     =   "Busca ORDENES pendientes de Pago"
         Top             =   930
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdPendientesFarmacia 
         Caption         =   "..."
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
         Left            =   2850
         TabIndex        =   77
         ToolTipText     =   "Busca PREVENTAS pendientes de Pago"
         Top             =   1290
         Visible         =   0   'False
         Width           =   495
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
         Height          =   2355
         Left            =   6360
         TabIndex        =   18
         Top             =   330
         Width           =   6600
         Begin VB.CommandButton cmbBuscaReceta 
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
            Left            =   4980
            Picture         =   "ucGestionCaja.ctx":834C
            Style           =   1  'Graphical
            TabIndex        =   181
            Top             =   1950
            Width           =   345
         End
         Begin VB.CommandButton btnBuscaCtaPorApellidos 
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
            Left            =   4935
            Picture         =   "ucGestionCaja.ctx":88D6
            Style           =   1  'Graphical
            TabIndex        =   180
            Top             =   405
            Width           =   345
         End
         Begin VB.TextBox txtNreceta 
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
            Left            =   3525
            MaxLength       =   10
            TabIndex        =   93
            Top             =   1965
            Width           =   1440
         End
         Begin VB.TextBox txtDni 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2655
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   91
            Top             =   420
            Width           =   1125
         End
         Begin VB.CommandButton cmdPaquetes 
            Caption         =   "Carga Paquetes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   5595
            TabIndex        =   62
            Top             =   1500
            Width           =   885
         End
         Begin VB.ComboBox cmbIdTipoFinanciamiento 
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
            Left            =   135
            TabIndex        =   60
            Top             =   1965
            Width           =   3360
         End
         Begin VB.TextBox txtNdocumentoB 
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
            Left            =   3465
            MaxLength       =   8
            TabIndex        =   50
            Top             =   1185
            Width           =   1830
         End
         Begin VB.TextBox txtNserieB 
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
            Left            =   2715
            MaxLength       =   4
            TabIndex        =   49
            Top             =   1185
            Width           =   675
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
            Height          =   330
            Left            =   3795
            MaxLength       =   10
            TabIndex        =   47
            Top             =   420
            Width           =   1110
         End
         Begin VB.ComboBox cmbOrdenes 
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
            Left            =   135
            Locked          =   -1  'True
            TabIndex        =   4
            Text            =   "cmbOrdenes"
            Top             =   1185
            Width           =   2535
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
            Left            =   135
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   420
            Width           =   1515
         End
         Begin VB.TextBox txtNroHistoria 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1635
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   3
            Top             =   420
            Width           =   1005
         End
         Begin VB.CommandButton cmdLeer 
            Caption         =   "Buscar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   5610
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   780
            Width           =   885
         End
         Begin VB.Label lblBuscaDNIReniec 
            Caption         =   "(Reniec)"
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
            Height          =   255
            Left            =   3015
            TabIndex        =   172
            Top             =   180
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lblNHistoriDNI 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo historia             Nro Historia     DNI                     Nro Cuenta          "
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
            Left            =   135
            TabIndex        =   171
            Top             =   180
            Width           =   5205
         End
         Begin VB.Label lblCuentaConSeguro 
            Caption         =   "....."
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
            Height          =   315
            Left            =   5325
            TabIndex        =   94
            Top             =   420
            Width           =   1215
         End
         Begin VB.Label lblOrden 
            BackStyle       =   0  'Transparent
            Caption         =   "N° Orden                                           N° Serie    N° Documento                       "
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
            Left            =   195
            TabIndex        =   36
            Top             =   915
            Width           =   5265
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Producto/Plan                                                      N° Receta"
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
            Left            =   135
            TabIndex        =   61
            Top             =   1755
            Width           =   5355
         End
      End
      Begin VB.ComboBox cmbIdCaja 
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
         Left            =   5670
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   1050
         Width           =   1905
      End
      Begin VB.ComboBox cmbIdTurno 
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
         Left            =   7620
         Style           =   2  'Dropdown List
         TabIndex        =   67
         Top             =   1050
         Width           =   1365
      End
      Begin VB.ComboBox cmbIdPuntoDeCarga 
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
         Left            =   9030
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   1050
         Width           =   2115
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
         Left            =   11190
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   1050
         Width           =   1005
      End
      Begin VB.Frame frmPreventaServ 
         Height          =   915
         Left            =   1770
         TabIndex        =   57
         Top             =   8340
         Visible         =   0   'False
         Width           =   2595
         Begin VB.CommandButton btnGeneraCuenta 
            Caption         =   "Genera N° Cuenta"
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
            Left            =   1320
            TabIndex        =   59
            Top             =   270
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.CheckBox chkGeneraPreventaServ 
            Caption         =   "Genera PreVenta"
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
            Left            =   60
            TabIndex        =   58
            Top             =   210
            Width           =   1155
         End
      End
      Begin SISGalenPlus.UcFacturacionContado UcFacturacionContado1 
         Height          =   2865
         Left            =   180
         TabIndex        =   48
         Top             =   4860
         Visible         =   0   'False
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   5054
      End
      Begin SISGalenPlus.ucFacturacionItems ucFacturacionProductos 
         Height          =   3540
         Left            =   150
         TabIndex        =   38
         Top             =   4200
         Width           =   12855
         _ExtentX        =   22781
         _ExtentY        =   7303
      End
      Begin VB.Frame fraOpciones 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   4560
         TabIndex        =   37
         Top             =   60
         Width           =   4125
         Begin Threed.SSOption optFarmacia 
            Height          =   285
            Left            =   4620
            TabIndex        =   1
            Top             =   60
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   503
            _Version        =   262144
            PictureFrames   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "ucGestionCaja.ctx":8E60
         End
         Begin Threed.SSOption optServicios 
            Height          =   345
            Left            =   840
            TabIndex        =   0
            Top             =   0
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   609
            _Version        =   262144
            PictureFrames   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "ucGestionCaja.ctx":9434
            Value           =   -1
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "CAJA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   4530
         TabIndex        =   30
         Top             =   8370
         Width           =   8475
         Begin VB.TextBox txtVuelto 
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
            Left            =   1530
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   225
            Width           =   1170
         End
         Begin VB.TextBox txtFalta 
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
            Left            =   4230
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   225
            Width           =   1095
         End
         Begin VB.TextBox txtEfectivo 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   6900
            TabIndex        =   8
            Top             =   225
            Width           =   1395
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "VUELTO (S/.)"
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
            Left            =   570
            TabIndex        =   35
            Top             =   300
            Width           =   945
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "FALTA (S/.)"
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
            Left            =   3390
            TabIndex        =   34
            Top             =   300
            Width           =   840
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "EFECTIVO(S/.)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5520
            TabIndex        =   33
            Top             =   300
            Width           =   1350
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "TOTALES"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   150
         TabIndex        =   19
         Top             =   7740
         Width           =   12870
         Begin VB.TextBox txtServicioSocial 
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
            Left            =   930
            MaxLength       =   30
            TabIndex        =   107
            Top             =   270
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.ComboBox cmbServicioSocial 
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
            Left            =   1560
            TabIndex        =   92
            Top             =   270
            Visible         =   0   'False
            Width           =   3210
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   11250
            TabIndex        =   28
            Top             =   210
            Width           =   1425
         End
         Begin VB.TextBox txtPagoACuenta 
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
            Height          =   405
            Left            =   8610
            TabIndex        =   27
            Top             =   210
            Width           =   1080
         End
         Begin VB.TextBox txtExonerado 
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
            Height          =   405
            Left            =   5910
            TabIndex        =   25
            Top             =   210
            Width           =   1170
         End
         Begin VB.TextBox txtPendientePago 
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
            Height          =   285
            Left            =   2670
            TabIndex        =   23
            Top             =   90
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.TextBox txtIngresado 
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
            Height          =   285
            Left            =   2490
            TabIndex        =   21
            Top             =   90
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.Label lblServicioSocial 
            AutoSize        =   -1  'True
            Caption         =   "Serv.Social"
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
            Left            =   90
            TabIndex        =   106
            Top             =   330
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TOTAL(S/.) "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   10125
            TabIndex        =   29
            Top             =   300
            Width           =   1140
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   7365
            TabIndex        =   26
            Top             =   330
            Width           =   1215
         End
         Begin VB.Label ll 
            AutoSize        =   -1  'True
            Caption         =   "EXONERADO"
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
            Left            =   4950
            TabIndex        =   24
            Top             =   330
            Width           =   930
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "PEND. PAGO"
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
            Left            =   2505
            TabIndex        =   22
            Top             =   330
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "INGRESADO"
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
            Left            =   2070
            TabIndex        =   20
            Top             =   120
            Visible         =   0   'False
            Width           =   405
         End
      End
      Begin VB.Frame Frame3 
         Height          =   915
         Left            =   150
         TabIndex        =   17
         Top             =   8340
         Width           =   1590
         Begin VB.CommandButton btnAceptar 
            Caption         =   "Aceptar (F2)"
            DisabledPicture =   "ucGestionCaja.ctx":9A53
            DownPicture     =   "ucGestionCaja.ctx":9EB3
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   90
            Picture         =   "ucGestionCaja.ctx":A328
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   180
            Width           =   1455
         End
      End
      Begin VB.Frame fraPaciente 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         Left            =   150
         TabIndex        =   14
         Top             =   2625
         Width           =   12840
         Begin VB.CommandButton cmdProcesaHistoricos 
            Caption         =   "..."
            Height          =   315
            Left            =   12585
            TabIndex        =   186
            Top             =   1125
            Width           =   165
         End
         Begin VB.TextBox txtDireccionProv 
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
            Left            =   6510
            MaxLength       =   100
            TabIndex        =   175
            Top             =   780
            Width           =   6240
         End
         Begin VB.TextBox txtEmailProv 
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
            Left            =   1395
            MaxLength       =   100
            TabIndex        =   173
            Top             =   765
            Width           =   3015
         End
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
            Height          =   315
            Left            =   1380
            MaxLength       =   100
            TabIndex        =   63
            Top             =   1110
            Width           =   11175
         End
         Begin VB.TextBox txtFechaBoleta 
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
            Left            =   6525
            TabIndex        =   42
            Top             =   405
            Width           =   1485
         End
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
            Left            =   165
            MaxLength       =   100
            TabIndex        =   5
            Top             =   405
            Width           =   4260
         End
         Begin VB.TextBox txtNroSerie 
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
            Left            =   10380
            MaxLength       =   30
            TabIndex        =   51
            Top             =   405
            Width           =   675
         End
         Begin VB.TextBox txtNroDocumento 
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
            Left            =   11265
            MaxLength       =   30
            TabIndex        =   11
            Top             =   405
            Width           =   1470
         End
         Begin VB.ComboBox cmbIdTipoComprobante 
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
            Left            =   8430
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   390
            Width           =   1830
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
            Left            =   4455
            MaxLength       =   11
            TabIndex        =   6
            Top             =   405
            Width           =   1770
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "Dirección Proveedor"
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
            Left            =   5010
            TabIndex        =   176
            Top             =   825
            Width           =   1440
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Email Proveedor"
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
            Left            =   165
            TabIndex        =   174
            Top             =   825
            Width           =   1155
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones"
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
            Left            =   165
            TabIndex        =   64
            Top             =   1140
            Width           =   1065
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   11115
            TabIndex        =   16
            Top             =   405
            Width           =   105
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   $"ucGestionCaja.ctx":A79D
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
            Left            =   180
            TabIndex        =   15
            Top             =   180
            Width           =   12375
         End
      End
      Begin VB.TextBox txtNombres 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6090
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   70
         Top             =   1470
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtFechaApertura 
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
         Left            =   6750
         MaxLength       =   30
         TabIndex        =   71
         Top             =   1470
         Visible         =   0   'False
         Width           =   315
      End
      Begin Threed.SSOption optCobrarOrdenExistente 
         Height          =   255
         Left            =   180
         TabIndex        =   73
         Top             =   990
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   450
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "PAGAR ORDEN EXISTENTE (SERVICIOS)"
      End
      Begin Threed.SSOption optReimprimirComprobante 
         Height          =   255
         Left            =   3330
         TabIndex        =   74
         Top             =   1890
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   450
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "REIMPRIMIR COMPROBANTE"
      End
      Begin Threed.SSOption optRealizarDevolucion 
         Height          =   255
         Left            =   3330
         TabIndex        =   75
         Top             =   2160
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   450
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "DEVOLUCIÓN ORDEN DE PAGO"
      End
      Begin Threed.SSOption OptOrdenFarmacia 
         Height          =   255
         Left            =   180
         TabIndex        =   76
         Top             =   1365
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   450
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "PAGAR ORDEN EXISTENTE (FARMACIA)"
      End
      Begin VB.ComboBox cmbOrdenProvenienteDe 
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
         ItemData        =   "ucGestionCaja.ctx":A869
         Left            =   3780
         List            =   "ucGestionCaja.ctx":A873
         TabIndex        =   79
         Text            =   "Combo1"
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin Threed.SSOption optPagarEstadoDeCuenta 
         Height          =   255
         Left            =   210
         TabIndex        =   89
         Top             =   2085
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   450
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "PAGAR CUENTA TOTAL (SERVICIOS)"
      End
      Begin Threed.SSOption optPagarEstadoDeCTAFarmacia 
         Height          =   255
         Left            =   180
         TabIndex        =   90
         Top             =   1710
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "PAGAR CUENTA TOTAL (FARMACIA)"
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Caja                                    Turno                      Punto de Carga                       Fecha Ingreso"
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
         Left            =   5670
         TabIndex        =   69
         Top             =   840
         Width           =   6915
      End
      Begin VB.Label lblNombre 
         BackColor       =   &H00373842&
         Caption         =   "Gestion de Caja"
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
         Height          =   465
         Left            =   -74880
         TabIndex        =   13
         Top             =   390
         Width           =   12945
      End
   End
End
Attribute VB_Name = "ucGestionCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para mantenimiento de CAJA (Emisión de Boletas, Tickets,...)
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'----------------------------------------------------------------------------------------
Option Explicit
Dim mo_Formulario As New sighEntidades.Formulario
Dim mi_Opcion As sghopcionespago
Dim mo_Teclado As New sighEntidades.Teclado
Dim lcBuscaParametro As New SIGHDatos.Parametros
'
Dim mo_doCajaGestion As DOCajaGestion
Dim mo_DOCajaCaja As DOCajaCaja
Dim mo_DOComprobantePago As New DOCajaComprobantesPago
Dim mo_DoNotaCreditoDebito As New DoNotaCreditoDebito 'Frank 24082015
Dim mo_DOComprobantePagoDevolucion As New DOCajaComprobantesPago
Dim mo_oComprobantepago As New CajaComprobantesPago
Dim mo_DOFactOrdenServicio As New DOFactOrdenServicio
Dim mo_DOFactOrdenBienInsumo As New DoFactOrdenesBienes
Dim mo_DoAtencion As New DOAtencion
Dim mo_DoFactOrdenServPagos  As New DoFactOrdenServPagos
Dim mo_DOCuentaAtencion As DOCuentaAtencion
Dim doCajero As New SIGHComun.DOCajaCajero
Dim mo_DoPaciente As New doPaciente
Dim mo_Reniec As New SIGHNegocios.ReniecGalenhosNegocios '
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_AdminComun As New SIGHNegocios.ReglasComunes
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_sighProxies As New SIGHProxies.Procesos
Dim mo_cmbIdPuntoCarga As New sighEntidades.ListaDespleglable
Dim mo_cmbIdEstado As New sighEntidades.ListaDespleglable
Dim mo_cmbFechaIngreso As New sighEntidades.ListaDespleglable
Dim mo_cmbIdTipoGenHistoriaClinica As New sighEntidades.ListaDespleglable
Dim mo_cmbIdCaja As New ListaDespleglable
Dim mo_cmbIdTurno As New ListaDespleglable
Dim mo_cmbIdCajaBusqueda As New ListaDespleglable
Dim mo_cmbIdTurnoBusqueda As New ListaDespleglable
Dim mo_cmbIdTipoComprobante As New ListaDespleglable
Dim mo_cmbOrdenes As New ListaDespleglable
Dim mo_cmbIdResponsable As New sighEntidades.ListaDespleglable
Dim mo_cmbIdTipoFinanciamiento As New sighEntidades.ListaDespleglable
Dim mo_cmbServicioSocial As New sighEntidades.ListaDespleglable
'
Dim oRsBusquedaRecibos As New ADODB.Recordset
Dim oRsBusquedaDevNotaCredito As New ADODB.Recordset 'Frank 24082015
Dim oRsCajeros As New Recordset
'
Dim ml_IdOrdenDespacho As Long
Dim ml_IdPaciente As Long
Dim ml_IdTipoFinanciamiento As Long
Dim md_Total As Double
Dim md_Ingresado As Double
Dim md_PendientePago As Double
Dim md_PagoACuenta As Double
Dim md_Exonerado As Double
Dim ml_TipoProducto As Long
Dim ml_idUsuario As Long
Dim ml_PuntoCarga As Long
Dim ml_idOrden As Long
Dim ms_MensajeError As String
Dim mb_ExistenDatos As Boolean
Dim mo_Apariencia As New sighEntidades.GridInfragistic
Dim ml_idCuentaAtencion As Long
Const ID_TIPO_COMPROBANTE_FACTURA = 2
Dim ml_IdGestionCaja As Long
Dim lbEsDevolucion As Boolean, lbItemEsDevolucion As Boolean
Dim ml_NombreCajero As String
Dim lnParametrosImprimeBoleta As sghImpresion
Dim ml_IdFormaPago As Long
Dim ml_IdFarmacia As Long
Dim ml_idPreVenta As Long
Dim ms_Descripcion As String 'JR 1105
Dim lbCargaEstadoDeCuentaFarmacia As Boolean    'True=Carga CUENTA DE FARMACIA en CAJA Servicios, false=carga CUENTA DE SERVICIO en CAJA Servicios
Dim ml_idConfiguracionParaPreventa As Long
Dim lbBoletaDeServicios As Boolean
Dim lnTotalGrid As Double
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim lbEsUnaFactura As Boolean
Dim lbTienePermisoSoloParaBoletaFarmacia As Boolean
Dim lbTienePermisoReimprimeBoleta As Boolean
Dim lbTienePermisoExonerarPacExterno As Boolean
Dim ml_Estado As Integer
Dim lnIdFactPaquete As Long
Const lcEFE As String = "F"     'debb-16/02/2011
Dim lbCargaEstadoDeCuentaFS As Boolean  'debb-17/02/2011
Dim lbTienePermisoSoloParaBoletaServicio As Boolean
Dim lnIdReceta As Long, lbSeAperturoCAJA As Boolean
Dim lnCantidadItemsDeLaBoleta As Long, lbElNroItemsEsMenorAlMaximoDeBoleta As Boolean
Public Event GuardoComprobante(bGuardo As Boolean)
Const lcTituloNotaCredito As String = "CAJA - Notas de Crédito"
Dim lbTieneLicenciaParaNotaCreditoYsunat As Boolean
Dim lbTrabajaComoCajero As Boolean, lbFacturaSinIGV As Boolean
Dim lnIGV As Double
Dim lcDNIbuscado As String    'debb-05/12/2017
Dim lcVuelto As String
Dim lbEstaCajaUsaDescripcionLarga As Boolean, lnMontoIGV99 As Double, lbTieneCredito99 As Boolean
Dim lnIdPacienteDelDNIelegido As Long, lbLaDevolucionNCesAutomatica As Boolean
Dim AutomNC_lnIdCaja  As Long, AutomNC_ml_IdUsuario As Long, AutomNC_lnIdGestionCaja As Long, AutomNC_lnIdTurno As Long
Dim lbPuedeVerVistaPrevia As Boolean, lbUsaResumenDiarioSunat As Boolean
Dim wxParametro579 As String, ldHoy As Date
Dim lbEsUnaRecetaOtrosCpt As Boolean

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Property Let NombreCajero(lValue As String)
   ml_NombreCajero = lValue
End Property

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property

'********************************************************************************************
'                                   COMPROBANTE DE PAGO
'********************************************************************************************
Property Let PuntoCarga(lValue As Long)
    ml_PuntoCarga = lValue
End Property

Property Get PuntoCarga() As Long
    PuntoCarga = ml_PuntoCarga
End Property

Property Let idTipoFinanciamiento(lValue As Long)
    ml_IdTipoFinanciamiento = lValue
End Property

Property Get idTipoFinanciamiento() As Long
    idTipoFinanciamiento = ml_IdTipoFinanciamiento
End Property

Property Let IdOrden(lValue As Long)
    ml_idOrden = lValue
End Property

Property Get IdOrden() As Long
    IdOrden = ml_idOrden
End Property

Property Let IdGestionCaja(lValue As Long)
    ml_IdGestionCaja = lValue
End Property

Property Get IdGestionCaja() As Long
    IdGestionCaja = ml_IdGestionCaja
End Property

'********************************************************************************************
'                                   GESTION DE CAJA
'********************************************************************************************
Sub SkinConfigura()
  On Error GoTo ErrSkin
  If sighEntidades.Parametro282valorInt = "1" Then
        btnBuscar.Picture = LoadPicture(App.Path & "\Binoculr.ico")
        btnBuscar.Caption = ""
        btnLimpiar.Picture = LoadPicture(App.Path & "\Refresh.ico")
        btnLimpiar.Caption = ""
        mo_Apariencia.ConfigurarFilasBiColores grdGestionCaja, "99"
        mo_Apariencia.ConfigurarFilasBiColores grdNotaCredito, "99"
        lblNombre.Alignment = 2
        lblNombre.BackColor = vbBlue
  Else
        mo_Apariencia.ConfigurarFilasBiColores grdGestionCaja, sighEntidades.GrillaConFilasBicolor
        mo_Apariencia.ConfigurarFilasBiColores grdNotaCredito, sighEntidades.GrillaConFilasBicolor
  End If
ErrSkin:
End Sub

Public Function Inicializar()
    ldHoy = CDate(lcBuscaParametro.RetornaFechaServidorSQL)
    wxHuboCambioAfactura = False
    SkinConfigura
    Set mo_cmbIdCaja.MiComboBox = cmbIdCaja
    Set mo_cmbIdTurno.MiComboBox = cmbIdTurno

    Set mo_cmbIdCajaBusqueda.MiComboBox = cmbIdCajaBusqueda
    Set mo_cmbIdTurnoBusqueda.MiComboBox = cmbIdTurnoBusqueda

    Set mo_cmbIdPuntoCarga.MiComboBox = cmbIdPuntoDeCarga
    Set mo_cmbFechaIngreso.MiComboBox = cmbFechaIngreso
    Set mo_cmbIdTipoGenHistoriaClinica.MiComboBox = cmbIdTipoGenHistoriaClinica
    Set mo_cmbIdTipoComprobante.MiComboBox = cmbIdTipoComprobante
    Set mo_cmbOrdenes.MiComboBox = cmbOrdenes
    Set mo_cmbIdResponsable.MiComboBox = cmbIdResponsable
    Set mo_cmbIdTipoFinanciamiento.MiComboBox = cmbIdTipoFinanciamiento
    Set mo_cmbServicioSocial.MiComboBox = cmbServicioSocial
    
    txtFdesde.Text = Date & " 00:01"
    txtFhasta.Text = Date & " 23:59"
    
    txtFechaBoleta.Text = lcBuscaParametro.RetornaFechaServidorSQL()
    txtNotaFechaCanje.Text = lcBuscaParametro.RetornaFechaServidorSQL() 'Frank 24082015
    
    ConfigurarTurno
    ConfigurarCaja
    ConfigurarPuntosDeCarga
    ConfigurarTiposHistoriaClinica
    ConfigurarTipoComprobante
    ConfigurarSiSeImprimeBoleta
    ConfiguraPermisos
    ConfigurarTipoFinanciamiento
    '
    mo_cmbServicioSocial.BoundColumn = "IdEmpleado"
    mo_cmbServicioSocial.ListField = "Empleado"
    Set mo_cmbServicioSocial.RowSource = mo_ReglasFacturacion.EmpleadosSeleccionarPorFiltro("Where idLaboraArea=" & sghAreasLaboraEmpleado.sghSeguros & " and idLaboraSubArea= 9")
    '
    mo_Formulario.HabilitarDeshabilitar cmbIdCaja, False
    mo_Formulario.HabilitarDeshabilitar cmbIdTurno, False
    mo_Formulario.HabilitarDeshabilitar txtFechaApertura, False
    
    mo_Formulario.HabilitarDeshabilitar cmbIdTipoGenHistoriaClinica, False
    mo_Formulario.HabilitarDeshabilitar txtNroHistoria, False
    mo_Formulario.HabilitarDeshabilitar txtNombres, False
    mo_Formulario.HabilitarDeshabilitar cmbOrdenes, False
    mo_Formulario.HabilitarDeshabilitar cmbIdPuntoDeCarga, False
    mo_Formulario.HabilitarDeshabilitar cmbFechaIngreso, False
    mo_Formulario.HabilitarDeshabilitar txtNroSerie, False
    mo_Formulario.HabilitarDeshabilitar txtNroDocumento, False
    mo_Formulario.HabilitarDeshabilitar txtObservaciones, False
    mo_Formulario.HabilitarDeshabilitar txtDni, False
    
    UserControl.tabGestionCaja.TabVisible(1) = False
    UserControl.tabGestionCaja.TabVisible(2) = False 'Frank 24082015
    
    ucFacturacionProductos.TipoProducto = sghServicio
    ml_TipoProducto = sghServicio
    
    ucFacturacionProductos.idUsuario = ml_idUsuario
    ucFacturacionProductos.Inicializar
    optNuevoOrdenPagoConHistoria.Value = True
    optNuevoOrdenPagoConHistoria_Click True
    
    'Para estados de cuenta
    tabFactProductosPorCuenta.Left = ucFacturacionProductos.Left
    tabFactProductosPorCuenta.Top = ucFacturacionProductos.Top
    tabFactProductosPorCuenta.Width = ucFacturacionProductos.Width
    tabFactProductosPorCuenta.Height = ucFacturacionProductos.Height
    
    ucFactServiciosPorCuenta.idUsuario = ml_idUsuario
    ucFactServiciosPorCuenta.InHabilitaEdicionColumnasDelGrid = True
    
    ucFactServiciosPorCuenta.Inicializar

    ucFactBienesPorCuenta.idUsuario = ml_idUsuario
    ucFactBienesPorCuenta.InHabilitaEdicionColumnasDelGrid = True
    ucFactBienesPorCuenta.Inicializar
    
    ml_IdGestionCaja = -1
    
    If optFarmacia.Value = True Then
       optPagarEstadoDeCuenta.Visible = False
       optPagarEstadoDeCTAFarmacia.Visible = False
    Else
       optPagarEstadoDeCuenta.Visible = True
       optPagarEstadoDeCTAFarmacia.Visible = True
    End If
    'Configuracion para PreVenta (FARMACIA)
    ml_idConfiguracionParaPreventa = Val(lcBuscaParametro.SeleccionaFilaParametro(229))
    '
    InicilizarParametros
    '''''''''''''''''''''''''''''''''''''''Frank 24082015
    mo_Formulario.HabilitarDeshabilitar txtSerieNota, False
    mo_Formulario.HabilitarDeshabilitar txtDocumentoNota, False
    mo_Formulario.HabilitarDeshabilitar txtEstadoNota, False
    mo_Formulario.HabilitarDeshabilitar txtNotaRazonSocial, False
    mo_Formulario.HabilitarDeshabilitar txtNotaRuc, False
    mo_Formulario.HabilitarDeshabilitar txtNotaFechaEmision, False
    mo_Formulario.HabilitarDeshabilitar txtNotaCredDebDireccion, False
    mo_Formulario.HabilitarDeshabilitar txtNotaMotivo, False
    mo_Formulario.HabilitarDeshabilitar txtNotaConcepto, False
    mo_Formulario.HabilitarDeshabilitar txtNotaTotal, False
    mo_Formulario.HabilitarDeshabilitar txtNotaCredAprueba, False
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'kike 2017
    Dim lcMensajeLicencia As String
    lbTieneLicenciaParaNotaCreditoYsunat = True
    If wxParametro377 <> "S" Or sighEntidades.Parametro573valorInt = 2 Then
       lbTieneLicenciaParaNotaCreditoYsunat = False
    End If
    
    

    
    On Error Resume Next
    UserControl.txtNroCuenta.SetFocus
End Function



Sub InicilizarParametros()
        wxParametro102 = lcBuscaParametro.SeleccionaFilaParametro(102)
        wxParametro205 = lcBuscaParametro.SeleccionaFilaParametro(205)
        wxParametro206 = lcBuscaParametro.SeleccionaFilaParametro(206)
        wxParametro208 = lcBuscaParametro.SeleccionaFilaParametro(208)
        wxParametro207 = lcBuscaParametro.SeleccionaFilaParametro(207)
        wxParametro211 = lcBuscaParametro.SeleccionaFilaParametro(211)
        wxParametro221 = lcBuscaParametro.SeleccionaFilaParametro(221)
        wxParametro208 = lcBuscaParametro.SeleccionaFilaParametro(208)
        wxParametro285 = lcBuscaParametro.SeleccionaFilaParametro(285)
        wxParametro280 = lcBuscaParametro.SeleccionaFilaParametro(280)
        wxParametro286 = lcBuscaParametro.SeleccionaFilaParametro(286)
        wxParametro288 = lcBuscaParametro.SeleccionaFilaParametro(288)
        wxParametro339 = lcBuscaParametro.SeleccionaFilaParametro(339)
        wxParametro346 = lcBuscaParametro.SeleccionaFilaParametro(346)
        wxParametro379 = lcBuscaParametro.SeleccionaFilaParametro(379) 'sunat
        wxParametro381 = lcBuscaParametro.SeleccionaFilaParametro(381)
        wxParametro386 = lcBuscaParametro.SeleccionaFilaParametro(386)
        wxParametro387 = lcBuscaParametro.SeleccionaFilaParametro(387)
        wxParametro377 = lcBuscaParametro.SeleccionaFilaParametro(377)
        wxParametro500 = UCase(lcBuscaParametro.SeleccionaFilaParametro(500))   'debb-18/05/2016
        wxParametro501 = lcBuscaParametro.SeleccionaFilaParametro(501)  'debb-18/05/2016
        lcParametro523 = lcBuscaParametro.SeleccionaFilaParametro(523)
        lcParametro524 = lcBuscaParametro.SeleccionaFilaParametro(524)
        wxParametro527 = lcBuscaParametro.SeleccionaFilaParametro(527)
        wxParametro533 = lcBuscaParametro.SeleccionaFilaParametro(533)
        wxParametro534 = lcBuscaParametro.SeleccionaFilaParametro(534)
        wxParametro538 = lcBuscaParametro.SeleccionaFilaParametro(538)
        wxParametro543 = lcBuscaParametro.SeleccionaFilaParametro(543)
        wxParametro548 = lcBuscaParametro.SeleccionaFilaParametro(548)
        wxParametro549 = lcBuscaParametro.SeleccionaFilaParametro(549)
        wxParametro557 = lcBuscaParametro.SeleccionaFilaParametro(557)
        wxParametro558 = lcBuscaParametro.SeleccionaFilaParametro(558)
        wxParametro579 = lcBuscaParametro.SeleccionaFilaParametro(579)
        lbUsaResumenDiarioSunat = IIf(lcBuscaParametro.SeleccionaFilaParametro(571) = "S", True, False)
        lnIGV = Val(lcBuscaParametro.SeleccionaFilaParametro(221))
End Sub

'***************daniel barrantes**************
'***************Retorna si se Imprime BOLETA o solo PANTALLA
'***************
Sub ConfigurarSiSeImprimeBoleta()
    lnParametrosImprimeBoleta = mo_ReglasComunes.ParametrosSeleccionarValorIntPorTipoYCodigo("INDICADOR", "IMPRIMIR_RECIBO")
End Sub

'SUNAT
Public Function RealizarAperturaDeCaja(lIdUsuario As Long, lIdCaja As Long, lIdTurno As Long, lbEmiteSoloServicios As Boolean) As Boolean
Dim oDOCajaGestion As DOCajaGestion
Dim bAperturaOK As Boolean

    bAperturaOK = False
    Set oDOCajaGestion = mo_AdminCaja.RetornaCajaAbierta(lIdUsuario, lIdCaja, lIdTurno)
    Set mo_DOCajaCaja = mo_AdminCaja.CajaSeleccionarPorId(lIdCaja)
    If oDOCajaGestion Is Nothing Then

        Set oDOCajaGestion = New DOCajaGestion
        oDOCajaGestion.IdCaja = lIdCaja
        oDOCajaGestion.IdCajero = lIdUsuario
        oDOCajaGestion.IdTurno = lIdTurno
        oDOCajaGestion.EstadoLote = "A"
        oDOCajaGestion.FechaApertura = lcBuscaParametro.RetornaFechaHoraServidorSQL     'Now
        oDOCajaGestion.IdUsuarioAuditoria = lIdUsuario
        oDOCajaGestion.TotalCobrado = 0
        
        If mo_AdminCaja.CajaGestionAgregar(oDOCajaGestion) Then
            Set mo_doCajaGestion = oDOCajaGestion
            bAperturaOK = True
        End If
    Else
        Set mo_doCajaGestion = oDOCajaGestion
        bAperturaOK = True
    End If
    
    If bAperturaOK Then
        mo_cmbIdTurno.BoundText = oDOCajaGestion.IdTurno
        mo_cmbIdCaja.BoundText = oDOCajaGestion.IdCaja
        txtFechaApertura = oDOCajaGestion.FechaApertura
        mo_cmbIdTipoComprobante.BoundText = wxIdTipoComprobanteDefault        'Servicio
        mo_cmbIdPuntoCarga.BoundText = 99   'Cajero
        UserControl.tabGestionCaja.TabVisible(1) = True
        PermisosAccesoNotaCredito 'Frank 24082015
        UserControl.tabGestionCaja.Tab = 1
        UserControl.KeyPreview = True   'debb-16/02/2011
        '
        lbEstaCajaUsaDescripcionLarga = False
        If Val(mo_DOCajaCaja.idPartida) > 0 Then
           lbEstaCajaUsaDescripcionLarga = True
        End If
        CajaUsaDescripcionLarga
        '
    End If
    
    RealizarAperturaDeCaja = bAperturaOK
    If lbEmiteSoloServicios Then
       optServicios.Value = True
       optServicios.Visible = True
       optFarmacia.Value = False
       optFarmacia.Visible = False
    Else
       optServicios.Value = False
       optServicios.Visible = False
       optFarmacia.Value = True
       optFarmacia.Visible = True
    End If
    If optFarmacia.Value = True Then
       optPagarEstadoDeCuenta.Visible = False
       optPagarEstadoDeCTAFarmacia.Visible = False
    Else
       optPagarEstadoDeCuenta.Visible = True
       optPagarEstadoDeCTAFarmacia.Visible = True
    End If
    On Error Resume Next
    UserControl.txtNroCuenta.SetFocus
End Function

Public Function RealizarCierreDeCaja() As Boolean
    
    RealizarCierreDeCaja = False
    mo_doCajaGestion.fechaCierre = lcBuscaParametro.RetornaFechaHoraServidorSQL    'Now
    mo_doCajaGestion.EstadoLote = "C"
    
    If mo_AdminCaja.CajaGestionModificar(mo_doCajaGestion) Then
        UserControl.tabGestionCaja.TabVisible(1) = False
        UserControl.tabGestionCaja.TabVisible(2) = False 'Frank 24082015
        RealizarCierreDeCaja = True
        UserControl.KeyPreview = False   'debb-16/02/2011
    End If
    
End Function




Private Sub btnAceptar_KeyUp(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub btnBuscaCtaPorApellidos_Click()
    If txtNroCuenta.Locked = True Then
       Exit Sub
    End If
    Dim oBusqueda As New SIGHNegocios.BuscaPacientes
    Dim oDOPaciente As New doPaciente
    Dim oConexion As New Connection
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oBusqueda.TipoFiltro = sghFiltrarTodos
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oConexion)
        If Not oDOPaciente Is Nothing Then

            ml_IdPaciente = oDOPaciente.idPaciente
            Dim oRsTmp As New Recordset
            Set oRsTmp = mo_ReglasFarmacia.FacturacionCuentasAtencionSeleccionarPorIdPaciente(ml_IdPaciente, oConexion, True)
            If oRsTmp.RecordCount > 0 Then
               txtNroCuenta.Text = oRsTmp.Fields!idCuentaAtencion
            End If
            oRsTmp.Close
            Set oRsTmp = Nothing
            txtNroCuenta_KeyPress 13
        End If
    End If
    Set oDOPaciente = Nothing
    Set oBusqueda = Nothing
    oConexion.Close
    Set oConexion = Nothing
End Sub

Private Sub btnBuscar_Click()
    If IsDate(UserControl.txtFdesde.Text) = False Then
       MsgBox "La FECHA INICIAL es vacia ó no tiene el formato correcto", vbInformation, "CAJA"
       Exit Sub
    End If
    If IsDate(UserControl.txtFhasta.Text) = False Then
       MsgBox "La FECHA FINAL es vacia ó no tiene el formato correcto", vbInformation, "CAJA"
       Exit Sub
    End If
    If CDate(UserControl.txtFdesde.Text) > CDate(UserControl.txtFhasta.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, "CAJA"
       Exit Sub
    End If
    MousePointer = 11
    RealizarBusqueda
    MousePointer = 1
End Sub

Sub LimpiarOpciones()
    lbEsUnaRecetaOtrosCpt = False
    lnIdPacienteDelDNIelegido = 0
    lcDNIbuscado = ""
    Set mo_DOFactOrdenServicio = Nothing
    Set mo_DoAtencion = Nothing
    Set mo_DOComprobantePago = Nothing
    Set ucFacturacionProductos.Atencion = Nothing
    mo_DOFactOrdenServicio.IdOrden = 0
    mo_DOFactOrdenBienInsumo.IdOrden = 0
    mo_cmbIdTipoGenHistoriaClinica.BoundText = wxParametro211
    txtNroHistoria = ""
    txtNombres = ""
    cmbOrdenes.Text = ""
    mo_cmbFechaIngreso.BoundText = 0
    txtIngresado = ""
    txtPendientePago = ""
    txtPagoACuenta = "0"
    txtExonerado = "0"
    md_Total = 0
    txtTotal.Text = "0"
    txtEfectivo = ""
    txtFalta = ""
    txtVuelto = ""
    txtRazonSocial = ""
    txtRuc = ""
    txtNserieB.Text = ""
    txtNdocumentoB.Text = ""
    txtEmailProv.Text = ""
    txtRuc.Tag = 0
    txtDireccionProv.Text = ""

    ml_IdPaciente = 0
    ml_IdFormaPago = 1          'Contado
    ml_IdFarmacia = 0           '1=Farmacia Principal,2=Farmacia Emergencia,0-otros
    ml_idPreVenta = 0
    ml_idCuentaAtencion = 0
    txtNreceta.Text = "": lnIdReceta = 0
    'debb-16/02/2011
    cmdOrdenExistenteFS.Visible = IIf(mi_Opcion = sghPagarOrdenExistenteFS Or mi_Opcion = sghPagarOrdenExistente Or mi_Opcion = sghPagarOrdenExistenteF, True, False)
    'debb-16/02/2011

    txtNroCuenta.Text = ""
    ml_IdOrdenDespacho = 0
    mo_Formulario.HabilitarDeshabilitar txtExonerado, False
    If lbTienePermisoExonerarPacExterno = True Then
        mo_Formulario.HabilitarDeshabilitar txtExonerado, True
    End If
    UcFacturacionContado1.Visible = False
    'frmPreventaServ.Visible = False
    chkGeneraPreventaServ.Value = 0
    mo_cmbIdTipoFinanciamiento.BoundText = "1"
    ConfiguracionParaPreVenta
    
    mo_Formulario.HabilitarDeshabilitar cmbIdTipoGenHistoriaClinica, False
    mo_Formulario.HabilitarDeshabilitar txtNroHistoria, False
    mo_Formulario.HabilitarDeshabilitar txtNombres, False
    mo_Formulario.HabilitarDeshabilitar cmbOrdenes, False
    mo_Formulario.HabilitarDeshabilitar cmbIdPuntoDeCarga, False
    mo_Formulario.HabilitarDeshabilitar cmbFechaIngreso, False
    mo_Formulario.HabilitarDeshabilitar txtNroCuenta, False
    mo_Formulario.HabilitarDeshabilitar txtNserieB, False
    mo_Formulario.HabilitarDeshabilitar txtNdocumentoB, False
    mo_Formulario.HabilitarDeshabilitar txtPagoACuenta, False
    mo_Formulario.HabilitarDeshabilitar cmbIdTipoFinanciamiento, False


    
    
    mo_Formulario.HabilitarDeshabilitar txtObservaciones, IIf(wxParametro386 = "S", True, False)
    
    mo_Formulario.HabilitarDeshabilitar txtDni, False    'debb-18/02/2011
    lblBuscaDNIReniec.Visible = False
    
    cmdPaquetes.Enabled = False: lnIdFactPaquete = 0
    lbItemEsDevolucion = False: txtObservaciones.Text = ""
    'debb-16/02/2011
    lbCargaEstadoDeCuentaFS = False
    UserControl.KeyPreview = True
    txtDni.Text = ""
    UserControl.txtCtaFarmExonerado.Text = ""
    UserControl.txtCtaFarmTfarmacia.Text = ""
    UserControl.txtCtaServExonerado.Text = ""
    UserControl.txtCtaServTservicio.Text = ""
    lblCuentaConSeguro.Caption = ""
    
    UserControl.txtRazonSocial.Text = ""
    UserControl.txtRuc.Text = ""
    UserControl.txtDireccionProv.Text = ""
    UserControl.txtEmailProv.Text = ""
    
    'debb-16/02/2011
    cmbServicioSocial.Visible = False: txtServicioSocial.Visible = False: lblServicioSocial.Visible = False
    txtNreceta.Enabled = True: cmbBuscaReceta.Enabled = True
    txtRazonSocial.Enabled = True
End Sub





Private Sub btnGeneraCuenta_Click()
'    Dim oFacGeneraCtaPacienteExterno As New FacGeneraCtaPacienteExterno
'    oFacGeneraCtaPacienteExterno.Opcion = 1
'    oFacGeneraCtaPacienteExterno.idPuntoCarga = 99     'cajero
'    oFacGeneraCtaPacienteExterno.idUsuario = ml_idUsuario
'    oFacGeneraCtaPacienteExterno.lcNombrePc = mo_lcNombrePc
'    oFacGeneraCtaPacienteExterno.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
'    oFacGeneraCtaPacienteExterno.Show 1
'    Set oFacGeneraCtaPacienteExterno = Nothing
End Sub

Private Sub btnLimpiar_Click()
     mo_cmbIdCajaBusqueda.BoundText = ""
     mo_cmbIdTurnoBusqueda.BoundText = ""
     txtNroSerieBusqueda = ""
     txtNroDocumentoBusqueda = ""
     txtNroHistoriaBusqueda = ""
     txtFhasta.Text = Date & " 23:59"
     txtFdesde.Text = Date & " 00:01"
     TxtRsocial.Text = ""
     mo_cmbIdResponsable.BoundText = ""
     chkSoloCredito.Value = 0
End Sub










Private Sub cmbBuscaReceta_Click()
    Dim oBusqueda As New SIGHNegocios.clBuscaReceta
    oBusqueda.idPuntoCarga = sghPtoCargaCaja
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
       txtNreceta.Text = oBusqueda.IdRecetaSeleccionada
       txtNreceta_LostFocus
    End If
    Set oBusqueda = Nothing

End Sub

Private Sub cmbIdCajaBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
     mo_Teclado.RealizarNavegacion KeyCode, cmbIdCajaBusqueda
     AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdTipoComprobante_Click()
Dim lIdTipoComprobante As Long
Dim oCajaNroDocumento As New DOCajaNroDocumento
Dim rsBuscaBoleta As Recordset, lbSigue As Boolean, lnLen As Integer

    txtRuc.Enabled = True
    txtNroSerie.Text = ""
    txtNroDocumento.Text = ""

    lIdTipoComprobante = Val(mo_cmbIdTipoComprobante.BoundText)
    txtEmailProv.Enabled = True
    lbFacturaSinIGV = True
    txtDireccionProv.Enabled = True
        
    If lIdTipoComprobante > 0 Then
        Set oCajaNroDocumento = mo_AdminCaja.NroDocumentoSeleccionarPorIdCajaYTipoComprobante(mo_doCajaGestion.IdCaja, lIdTipoComprobante)
        txtNroSerie.Text = Trim(oCajaNroDocumento.nroSerie)
        txtNroDocumento.Text = Right("00000000" & Trim(oCajaNroDocumento.nrodocumento), 8)
        'comprueba que no existe esa nueva Boleta
        lbSigue = True
        Do While lbSigue = True
           Set rsBuscaBoleta = mo_AdminCaja.CajaComprobantePagoPorSerieDocumento(txtNroSerie.Text, txtNroDocumento.Text)
           If rsBuscaBoleta.RecordCount = 0 Then
              lbSigue = False
           Else
'              lnLen = Len(txtNroDocumento.Text)
              txtNroDocumento.Text = Right("00000000" & Trim(Str(Val(txtNroDocumento.Text) + 1)), 8)  'lnLen)
           End If
        Loop
        '
        If lIdTipoComprobante <> ID_TIPO_COMPROBANTE_FACTURA Then
            If mi_Opcion = sghNuevoPagoSinHistoria Then
                mo_Formulario.HabilitarDeshabilitar txtDni, True
            End If
            mo_Formulario.HabilitarDeshabilitar txtRuc, False
            mo_Formulario.HabilitarDeshabilitar txtEmailProv, False
            mo_Formulario.HabilitarDeshabilitar txtDireccionProv, False
            txtRuc.Text = ""
            lbEsUnaFactura = False
            txtRuc.Tag = 0
        Else
            If mi_Opcion = sghNuevoPagoSinHistoria Then mo_Formulario.HabilitarDeshabilitar txtDni, False
            wxHuboCambioAfactura = True
            txtDni.Text = ""
            mo_Formulario.HabilitarDeshabilitar txtRuc, True
            mo_Formulario.HabilitarDeshabilitar txtEmailProv, True
            mo_Formulario.HabilitarDeshabilitar txtDireccionProv, True
            lbEsUnaFactura = True
            lbFacturaSinIGV = IIf(oCajaNroDocumento.FacturaSinIGV = True, True, False)
        End If
        
    End If
    Set oCajaNroDocumento = Nothing



'    wxIdTipoComprobanteDefault = lIdTipoComprobante
'    CargaSetup_Caja App.Path & "\archivos", wxIdTipoComprobanteDefault
If lIdTipoComprobante = 3 And wxParametro527 = "S" Then
   CargaSetup_Caja App.Path & "\archivos", 4, False
Else
    CargaSetup_Caja App.Path & "\archivos", lIdTipoComprobante, False
End If
End Sub










Private Sub cmbIdTipoComprobante_KeyUp(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdTipoFinanciamiento_Click()
    ucFacturacionProductos.idTipoFinanciamiento = Val(mo_cmbIdTipoFinanciamiento.BoundText)
End Sub

Private Sub cmbIdTipoFinanciamiento_KeyUp(KeyCode As Integer, Shift As Integer)
   AdministrarKeyPreview KeyCode
End Sub



Private Sub cmbIdTipoGenHistoriaClinica_KeyUp(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdTurnoBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTurnoBusqueda

End Sub

Private Sub cmbOrdenes_Click()
    cmbOrdenes.Text = mo_cmbOrdenes.BoundText
End Sub





Private Sub cmbOrdenes_KeyPress(KeyAscii As Integer)
    'debb-16/02/2011
    If KeyAscii = 13 And Len(Trim(cmbOrdenes.Text)) > 0 Then
       If Len(Trim(cmbOrdenes.Text)) > 9 Then
          MsgBox "El Nª Orden no puede exceder de 9 caracteres", vbInformation, "caja"
          cmbOrdenes.Text = ""
          Exit Sub
       End If
       Dim lcCmbOrdenes As String
       lcCmbOrdenes = cmbOrdenes.Text
       If UCase(Right(cmbOrdenes.Text, 1)) = lcEFE Then
          If lbTienePermisoSoloParaBoletaServicio = True Then
              MsgBox "Solo tiene permiso para emitir Boletas de Servicio" & Chr(13) & "verifique PERMISOS del ROL", vbInformation, "Caja"
              lcCmbOrdenes = ""
              cmbOrdenes.Text = ""
              Exit Sub
          End If
          'Frank
          'mo_cmbIdTipoComprobante.BoundText = wxIdTipoComprobante2    'kike 2017
          OptOrdenFarmacia_Click 1
       Else
          If lbTienePermisoSoloParaBoletaFarmacia = True Then
              MsgBox "Solo tiene permiso para emitir Boletas de Farmacia" & Chr(13) & "verifique PERMISOS del ROL", vbInformation, "Caja"
              lcCmbOrdenes = ""
              cmbOrdenes.Text = ""
              Exit Sub
          End If
          optCobrarOrdenExistente_Click 1
       End If
       cmbOrdenes.Text = lcCmbOrdenes
       cmdLeer_Click
    End If
    'debb-16/02/2011
End Sub

Private Sub cmbOrdenes_KeyUp(KeyCode As Integer, Shift As Integer)
     AdministrarKeyPreview KeyCode
End Sub



Private Sub cmbOrdenes_LostFocus()

    If Len(cmbOrdenes.Text) > 9 Then
       MsgBox "No puede pasar de 9 caracteres", vbInformation, "Caja"
       On Error Resume Next
       cmbOrdenes.Text = ""
       cmbOrdenes.SetFocus
    End If
End Sub



Private Sub cmbServicioSocial_KeyUp(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmdBuscarPorApell_Click()
    txtNroHistoriaBusqueda.Text = ""
    Dim oBusqueda As New SIGHNegocios.BuscaPacientes
    Dim oDOPaciente As New doPaciente
    Dim oConexion As New Connection
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oBusqueda.TipoFiltro = sghFiltrarTodos
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oConexion)
        If Not oDOPaciente Is Nothing Then
           txtNroHistoriaBusqueda.Text = oDOPaciente.NroHistoriaClinica
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oDOPaciente = Nothing
    Set oBusqueda = Nothing
    If txtNroHistoriaBusqueda.Text <> "" Then
       txtNroHistoriaBusqueda_KeyPress 13
    End If
End Sub

Private Sub cmdLeer_Click()
    Select Case mi_Opcion
    'En el caso de las cuenta se va a mostrar Servicios y Bienes e Insumos
    Case sghopcionespago.sghPagarCuentaExistente, sghopcionespago.sghPagarCuentaTotalFS   'debb-17/02/2011
        CargarDatosServiciosALosControlesPorIdCuentaAtencion ml_idCuentaAtencion
        On Error Resume Next
        btnAceptar.SetFocus
    Case sghopcionespago.sghNuevoPagoConHistoria
        ucFacturacionProductos.TabEnDescripcion
    Case Else
        If mi_Opcion = sghPagarOrdenExistenteF Then 'solo es usado cuando es "pagar orden existente"
            LeerBienesPorTipoDePago
        Else
            Select Case ml_TipoProducto
            Case sghServicio
                LeerServiciosPorTipoDePago
            Case sghbien
                LeerBienesPorTipoDePago
            End Select
        End If
    End Select
  
End Sub

Private Sub optCobrarCuentaExistente_Click()
    
    mi_Opcion = sghopcionespago.sghPagarCuentaExistente
    LimpiarOpciones
    ucFacturacionProductos.LimpiarGrilla
    cmdLeer.Visible = True
    mo_Formulario.HabilitarDeshabilitar cmbIdTipoGenHistoriaClinica, True
    mo_Formulario.HabilitarDeshabilitar txtNroHistoria, True
    mo_Formulario.HabilitarDeshabilitar txtNombres, False
    mo_Formulario.HabilitarDeshabilitar cmbOrdenes, False
    mo_Formulario.HabilitarDeshabilitar cmbIdPuntoDeCarga, False
    mo_Formulario.HabilitarDeshabilitar cmbFechaIngreso, False
    mo_Formulario.HabilitarDeshabilitar txtDni, False  'debb-18/02/2011
End Sub

'debb-16/02/2011
Private Sub cmdOrdenExistenteFS_Click()
    Dim oBusqueda As New OrdenesPendientesPagoBusqueda
    oBusqueda.TipoProducto = sghAmbos
    oBusqueda.Show 1
    If oBusqueda.BotonPresionado = sghAceptar Then
       cmbOrdenes.Text = oBusqueda.idOrdenSeleccionado
       If oBusqueda.lbEstoyEnTabServicio = False Then
          cmbOrdenes.Text = Trim(cmbOrdenes.Text) & lcEFE
       End If
       cmbOrdenes_KeyPress 13
    End If
    Set oBusqueda = Nothing
End Sub


Private Sub cmdOrdenExistenteFS_KeyUp(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmdPaquetes_Click()
    Dim oPaquetesBuscar As New SIGHNegocios.BuscaPaquetes
    oPaquetesBuscar.DebeConsiderarPaquete = sghTipoPaqueteSoloServicio
    oPaquetesBuscar.MostrarFormulario
    If oPaquetesBuscar.BotonPresionado = sghAceptar Then
       lnIdFactPaquete = oPaquetesBuscar.idFactPaquete
       ucFacturacionProductos.PaqueteServicioAgregaProductos lnIdFactPaquete
       Dim oRsTmp As New Recordset
       Set oRsTmp = mo_ReglasFacturacion.FactCatalogoPaqueteSeleccionarPorFiltro(" idFactPaquete=" & lnIdFactPaquete)
       If oRsTmp.RecordCount > 0 Then
          txtRazonSocial.Text = oRsTmp.Fields!Codigo & " " & Trim(oRsTmp.Fields!descripcion)
          txtObservaciones.Text = "Pqte: " & Trim(oRsTmp.Fields!descripcion)
          mo_cmbIdTipoFinanciamiento.BoundText = oRsTmp.Fields!idTipoFinanciamiento
       End If
       oRsTmp.Close
       Set oRsTmp = Nothing
    End If
    Set oPaquetesBuscar = Nothing
End Sub

Private Sub cmdPendientesFarmacia_Click()
    Dim oBusqueda As New OrdenesPendientesPagoBusqueda
    oBusqueda.TipoProducto = sghbien
    oBusqueda.Show 1
    If oBusqueda.BotonPresionado = sghAceptar Then
       cmbOrdenes.Text = oBusqueda.idOrdenSeleccionado
       cmdLeer_Click
    End If
    Set oBusqueda = Nothing

End Sub

Private Sub cmdPendientesFarmacia_KeyUp(KeyCode As Integer, Shift As Integer)
AdministrarKeyPreview KeyCode
End Sub

Private Sub cmdPendientesPago_Click()
    Dim oBusqueda As New OrdenesPendientesPagoBusqueda
    oBusqueda.TipoProducto = ml_TipoProducto
    oBusqueda.Show 1
    If oBusqueda.BotonPresionado = sghAceptar Then
       cmbOrdenes.Text = oBusqueda.idOrdenSeleccionado
       cmdLeer_Click
    End If
    Set oBusqueda = Nothing
End Sub



Private Sub cmdPendientesPago_KeyUp(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub


Private Sub cmdProcesaHistoricos_Click()
 If MsgBox("Debe hacer en una PC que no sea CAJA" & Chr(13) & _
            "c:\excel.xls, libro=Report, filaDeInicio=2," & Chr(13) & _
            "                   a=fechaBoleta(dd/mm/aaaa), c=serie,d=numero,e=razonSocial,f=codigoCpt/codigoSismed," & Chr(13) & _
            "                   h=cantidad,i=importe, l=dni," & Chr(13) & _
            "Esta seguro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
            Dim lcRpta As String
            lcRpta = UCase(InputBox("ingrese clave", ""))
            If UCase(lcRpta) = "MDB" Or (Val(Left(lcRpta, 2)) = Month(Date) And Val(Right(lcRpta, 2)) = Day(Date)) Then
               On Error GoTo ErrExc
               Dim lcError As String
               lcError = ""
               'llenar temporal de boletas
               Dim oRsBoletas As New Recordset
                With oRsBoletas
                      .Fields.Append "fecha", adDate
                      .Fields.Append "Serie", adVarChar, 4
                      .Fields.Append "Numero", adVarChar, 12
                      .Fields.Append "RazonS", adVarChar, 150
                      .Fields.Append "codigo", adVarChar, 12
                      .Fields.Append "cantidad", adInteger
                      .Fields.Append "Importe", adDouble
                      .Fields.Append "dni", adVarChar, 12, adFldIsNullable
                      .Fields.Append "tipo", adVarChar, 1
                      .Fields.Append "idPaciente", adInteger
                      .Fields.Append "ImporteBoleta", adDouble
                      .Fields.Append "Exoneraciones", adDouble
                      .Fields.Append "esMDB", adVarChar, 1
                      .LockType = adLockOptimistic
                      .Open
                End With
              
              Dim lnPrecioSIS As Double, lnPrecioSOAT As Double, lnPrecioConvenio As Double, lnPrecioESSSALUD As Double, lcEsMDB As String
              Dim oCommand As New ADODB.Command
              Dim oParameter As ADODB.Parameter
              Dim oConexion As New ADODB.Connection
              Dim oRsPaciente As New Recordset
              Dim oRs As New ADODB.Recordset
              Dim oRsCatalogo As New Recordset
              Dim oDllFactUCGestionCaja As New SighFacturacion.dllFactUCGestionCaja
              Dim EXL As Excel.Application
              Set EXL = New Excel.Application
              Dim W As Excel.Workbook
              'Set W = EXL.Workbooks.Open("c:\excel.xls")
              Dim s As Excel.Worksheet
              'Set s = W.Sheets("Report")
              Dim lnFor As Integer, lnFila As Integer, lcRango As String, lnFilaFinal As Integer, oRsTmp As New Recordset, lnIdCpt As Long, lcSql As String, lcCodigo As String
              Dim lcCPTcorta As String, lnPrecioPagante As Double, lnIdProducto As Long, lbContinuar As Boolean, lnCantidad As Long, lcDNI As String
              Dim lbCont2 As Boolean, lcRazons As String, Lctipo As String, lnImporte As Double, lcFecha As String, lcSerie As String, lcNumero As String
              Dim lnPrecio As Double, lcError1 As String, lnIdPaciente1 As Long, lnImporteMDB As Double
              If UCase(lcRpta) = "MDB" Then
                 '******************jala datos del Acces*****************
                 Dim oRsParametros As New Recordset
                 Dim oRsMDB As New Recordset
                 Dim oRsMDB1 As New Recordset
                 Dim oRsMDB2 As New Recordset
                 Dim oConexionMDB As New Connection
                 lcSql = "select ValorTexto from Parametros where idparametro=581"
                 oRsParametros.Open lcSql, sighEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
                 oConexionMDB.CommandTimeout = 900
                 oConexionMDB.CursorLocation = adUseClient
                 oConexionMDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source =" & Trim(oRsParametros!ValorTexto) & "\parametros.mdb"
                 lcSql = "select * from cajacomprobantesPago"
                 If oRsMDB.State = 1 Then oRsMDB.Close
                 oRsMDB.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
                 If oRsMDB.RecordCount > 0 Then
                    oRsMDB.MoveFirst
                    Do While Not oRsMDB.EOF
                        If oRsMDB!IdTipoOrden <> 1 Then
                            'farmacia
                            oRsBoletas.AddNew
                            oRsBoletas!fecha = oRsMDB!FechaCobranza
                            oRsBoletas!serie = oRsMDB!nroSerie
                            oRsBoletas!numero = oRsMDB!nrodocumento
                            oRsBoletas!razonS = Left(oRsMDB!razonSocial, 150)
                            oRsBoletas!Codigo = "FARMACIA"
                            oRsBoletas!Cantidad = Round(oRsMDB!Total, 0)
                            oRsBoletas!Importe = oRsMDB!Total
                            oRsBoletas!DNI = ""
                            oRsBoletas!tipo = "F"
                            oRsBoletas!idPaciente = IIf(IsNull(oRsMDB!idPaciente), 0, oRsMDB!idPaciente)
                            oRsBoletas!ImporteBoleta = oRsMDB!Total
                            oRsBoletas!exoneraciones = oRsMDB!exoneraciones
                            oRsBoletas!esMDB = "S"
                            oRsBoletas.Update
                        Else
                            'servicios
                            lcSql = "select idOrdenPago from FactOrdenServicioPagos where idComprobantePago=" & _
                                    oRsMDB!IdComprobantePago
                            If oRsMDB1.State = 1 Then oRsMDB1.Close
                            oRsMDB1.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
                            If oRsMDB1.RecordCount > 0 Then
                               oRsMDB1.MoveFirst
                               Do While Not oRsMDB1.EOF
                                  lcSql = "select * from facturacionServicioPagos where idOrdenPago=" & oRsMDB1!IdOrdenPago
                                  If oRsMDB2.State = 1 Then oRsMDB2.Close
                                  oRsMDB2.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
                                  If oRsMDB2.RecordCount > 0 Then
                                     oRsMDB2.MoveFirst
                                     Do While Not oRsMDB2.EOF
                                        lcSql = "select codigo from FactCatalogoServicios where idProducto=" & oRsMDB2!idProducto
                                        If oRsCatalogo.State = 1 Then oRsCatalogo.Close
                                        oRsCatalogo.Open lcSql, sighEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
                                        If oRsCatalogo.RecordCount = 0 Then
                                            lcError1 = lcError1 & "La boleta " & oRsMDB!nroSerie & " -" & oRsMDB!nrodocumento & " no existe codigo idProducto: " & oRsMDB2!idProducto & Chr(13) & Chr(10)
                                        Else
                                            oRsBoletas.AddNew
                                            oRsBoletas!fecha = oRsMDB!FechaCobranza
                                            oRsBoletas!serie = oRsMDB!nroSerie
                                            oRsBoletas!numero = oRsMDB!nrodocumento
                                            oRsBoletas!razonS = Left(oRsMDB!razonSocial, 150)
                                            oRsBoletas!Codigo = Left(Trim(oRsCatalogo!Codigo), 12)
                                            oRsBoletas!Cantidad = oRsMDB2!Cantidad
                                            oRsBoletas!Importe = oRsMDB2!Total
                                            oRsBoletas!DNI = ""
                                            oRsBoletas!tipo = "S"
                                            oRsBoletas!idPaciente = IIf(IsNull(oRsMDB!idPaciente), 0, oRsMDB!idPaciente)
                                            oRsBoletas!ImporteBoleta = oRsMDB!Total
                                            oRsBoletas!exoneraciones = oRsMDB!exoneraciones
                                            oRsBoletas!esMDB = "S"
                                            oRsBoletas.Update
                                        End If
                                        oRsMDB2.MoveNext
                                     Loop
                                  End If
                                  oRsMDB1.MoveNext
                               Loop
                            End If
                        End If
                        oRsMDB.MoveNext
                    Loop
                 End If
                 
                 oConexionMDB.Close
                 Set oRsParametros = Nothing
                 Set oRsMDB = Nothing
                 Set oRsMDB1 = Nothing
                 Set oRsMDB2 = Nothing
                 Set oConexionMDB = Nothing
                 
              Else
                '******************jala datos del Excel*****************
                lnFila = 2
                lnFilaFinal = 15000
                lcError = "Pasó abrir EXCEL "
                Set W = EXL.Workbooks.Open("c:\excel.xls")
                Set s = W.Sheets("Report")
                For lnFor = lnFila To lnFilaFinal
                    lcRango = "A" + Trim(Str(lnFor))
                    lcFecha = Trim(s.range(lcRango).Value)
                    lcRango = "C" + Trim(Str(lnFor))
                    lcSerie = s.range(lcRango).Value
                    lcRango = "D" + Trim(Str(lnFor))
                    lcNumero = s.range(lcRango).Value
                    lcRango = "E" + Trim(Str(lnFor))
                    lcRazons = s.range(lcRango).Value
                    lcRango = "F" + Trim(Str(lnFor))
                    lcCodigo = s.range(lcRango).Value
                    lcRango = "H" + Trim(Str(lnFor))
                    lnCantidad = Val(s.range(lcRango).Value)
                    lcRango = "I" + Trim(Str(lnFor))
                    lnImporte = Val(s.range(lcRango).Value)
                    lcRango = "L" + Trim(Str(lnFor))
                    lcDNI = s.range(lcRango).Value
                    lcRango = "M" + Trim(Str(lnFor))
                    Lctipo = s.range(lcRango).Value
                    If Len(Trim(lcFecha)) = 0 Then
                       Exit For
                    End If
                    lcError = "Pasó cargar linea: " & lnFor
                    oRsBoletas.AddNew
                    oRsBoletas!fecha = CDate(lcFecha)
                    oRsBoletas!serie = Trim(lcSerie)
                    oRsBoletas!numero = Left(Trim(lcNumero), 12)
                    oRsBoletas!razonS = Left(Trim(lcRazons), 150)
                    oRsBoletas!Codigo = Left(Trim(lcCodigo), 12)
                    oRsBoletas!Cantidad = lnCantidad
                    oRsBoletas!Importe = lnImporte
                    oRsBoletas!DNI = Left(Trim(lcDNI), 8)
                    oRsBoletas!tipo = Left(Lctipo, 1)
                    oRsBoletas!idPaciente = 0
                    oRsBoletas!ImporteBoleta = 0
                    oRsBoletas!exoneraciones = 0
                    oRsBoletas!esMDB = "N"
                    oRsBoletas.Update
                    lcError = "Pasó cargar Temporal linea: " & lnFor
                Next
              End If
              Set s = Nothing
             ' W.Save
              If lcRpta <> "MDB" Then
              W.Close
              End If
              Set W = Nothing
              Set EXL = Nothing
              'carga temporales y objeto de datos
              oConexion.CursorLocation = adUseClient
              oConexion.CommandTimeout = 300
              oConexion.Open sighEntidades.CadenaConexion
              
              
              
              lcError1 = ""
              oRsBoletas.Sort = "serie,numero"
              oRsBoletas.MoveFirst
              Do While Not oRsBoletas.EOF
                    If oRs.State = 1 Then
                       Set oRs = Nothing
                    End If
                    With oRs
                            .Fields.Append "IdFacturacionProducto", adInteger
                            .Fields.Append "IdProducto", adInteger
                            .Fields.Append "Codigo", adVarChar, 255, adFldIsNullable
                            .Fields.Append "NombreProducto", adVarChar, 255, adFldIsNullable
                            'mgaray201411a
                            .Fields.Append "labConfHIS", adVarChar, 3, adFldIsNullable
                            .Fields.Append "IdTipoFinanciamiento", adInteger
                            .Fields.Append "IdFuenteFinanciamiento", adInteger, , adFldIsNullable
                            .Fields.Append "Poliza", adVarChar, 255
                            .Fields.Append "TipoFinanciamiento", adVarChar, 255
                            .Fields.Append "Cantidad", adInteger
                            .Fields.Append "PrecioUnitario", adCurrency
                            .Fields.Append "TotalPorPagar", adCurrency
                            .Fields.Append "IdEstadoFacturacion", adInteger
                            .Fields.Append "IdPuntoCarga", adInteger
                            .Fields.Append "IdAtencion", adInteger, , adFldIsNullable
                            .Fields.Append "IdCajero", adInteger, , adFldIsNullable
                            .Fields.Append "FechaAutorizaPendiente", adDBTimeStamp, , adFldIsNullable
                            .Fields.Append "FechaAutorizaSeguro", adDBTimeStamp, , adFldIsNullable
                            .Fields.Append "FechaAutorizaDevolucion", adDBTimeStamp, , adFldIsNullable
                            .Fields.Append "FechaCajero", adDBTimeStamp, , adFldIsNullable
                            .Fields.Append "IdUsuarioAutorizaPendiente", adInteger, , adFldIsNullable
                            .Fields.Append "IdUsuarioAutorizaSeguro", adInteger, , adFldIsNullable
                            .Fields.Append "IdUsuarioAutorizaDevolucion", adInteger, , adFldIsNullable
                            .Fields.Append "IdServicioInternamiento", adInteger, , adFldIsNullable
                            .Fields.Append "IdUsuarioAuditoria", adInteger, , adFldIsNullable
                            .Fields.Append "EstadoLocal", adVarChar, 1
                            .Fields.Append "IdComprobantePago", adInteger, , adFldIsNullable
                            .Fields.Append "IdComprobantePagoDevolucion", adInteger, , adFldIsNullable
                            .Fields.Append "IdOrden", adInteger
                            .Fields.Append "movTipo", adVarChar, 1, adFldIsNullable
                            .Fields.Append "movNumero", adVarChar, 9, adFldIsNullable
                            .Fields.Append "SeUsaSinPrecio", adBoolean
                            .Fields.Append "PermiteEditarPrecio", adBoolean
                            .Fields.Append "PqteIdFactPaquete", adInteger
                            .Fields.Append "PqteIdPuntoCarga", adInteger
                            .Fields.Append "PqteIdEspecialidadServicio", adInteger
                            .Fields.Append "PqteGrupo", adInteger
                            .Fields.Append "CantidadSinEditar", adInteger
                            .CursorType = adOpenDynamic
                            .LockType = adLockOptimistic
                            .Open
                   End With
                   lnIdPaciente1 = 0
                   
                   If Val(oRsBoletas!DNI) > 0 Then
                        If oRsPaciente.State = 1 Then oRsPaciente.Close
                        oRsPaciente.Open "select idPaciente from Pacientes where nroDocumento='" & oRsBoletas!DNI & "'", oConexion, adOpenKeyset, adLockOptimistic
                        If oRsPaciente.RecordCount > 0 Then
                           lnIdPaciente1 = oRsPaciente!idPaciente
                        End If
                   ElseIf oRsBoletas!idPaciente > 0 Then
                        lnIdPaciente1 = oRsBoletas!idPaciente
                   End If
                   lcEsMDB = oRsBoletas!esMDB
                   txtNroSerie.Text = oRsBoletas!serie
                   txtNroDocumento.Text = oRsBoletas!numero
                   txtRazonSocial.Text = oRsBoletas!razonS
                   If lcEsMDB = "S" Then
                      txtFechaBoleta.Text = oRsBoletas!fecha
                   Else
                      txtFechaBoleta.Text = oRsBoletas!fecha & " " & Format(Now, "hh:mm:ss")
                   End If
                   txtExonerado.Text = oRsBoletas!exoneraciones
                   lcSerie = oRsBoletas!serie
                   lcNumero = oRsBoletas!numero
                   Lctipo = oRsBoletas!tipo
                   lnImporteMDB = oRsBoletas!ImporteBoleta
                   
                   lcError1 = ""
                   lnCantidad = 0
                   lnImporte = 0
                   Do While Not oRsBoletas.EOF And lcSerie = oRsBoletas!serie And lcNumero = oRsBoletas!numero
                        If Lctipo = "F" Then
                               lnImporte = lnImporte + oRsBoletas!Importe
                               lnCantidad = lnCantidad + oRsBoletas!Cantidad
                        Else
                                If oRsCatalogo.State = 1 Then oRsCatalogo.Close
                                oRsCatalogo.Open "select idProducto from FactCatalogoServicios where codigo='" & oRsBoletas!Codigo & "'", oConexion, adOpenKeyset, adLockOptimistic
                                If oRsCatalogo.RecordCount = 0 Then
                                   lcError1 = lcError1 & "La boleta " & lcSerie & " -" & lcNumero & " no existe codigo cpt: " & oRsBoletas!Codigo & Chr(13) & Chr(10)
                                Else
                                    lnPrecio = Round(oRsBoletas!Importe / oRsBoletas!Cantidad, 2)
                                    lnImporte = lnImporte + oRsBoletas!Importe
                                    oRs.AddNew
                                    oRs!Cantidad = oRsBoletas!Cantidad
                                    oRs!idProducto = oRsCatalogo!idProducto
                                    oRs!PrecioUnitario = lnPrecio
                                    oRs.Update
                                End If
                        End If
                        oRsBoletas.MoveNext
                        If oRsBoletas.EOF Then
                          Exit Do
                        End If
                   Loop
                   If lcEsMDB = "S" Then
                      lnImporte = lnImporteMDB
                   End If
                   If Lctipo = "F" Then
                            oRs.AddNew
                            oRs!Cantidad = lnCantidad
                            oRs!idProducto = 5961         ' 54253      'cpt=FARMACIA   codigo=000004
                            oRs!PrecioUnitario = Round(lnImporte / lnCantidad, 2)
                            oRs.Update
                            
                   End If
                   'carga datos a objetos
                  
                   mi_Opcion = sghNuevoPagoSinHistoria
                   CargaDatosAlObjetosDeDatos
                   mo_DoFactOrdenServPagos.fechacreacion = CDate(txtFechaBoleta.Text)
                   mo_DOComprobantePago.Total = lnImporte
                   mo_DOComprobantePago.FechaCobranza = CDate(txtFechaBoleta.Text)
                   mo_DOComprobantePago.idPaciente = lnIdPaciente1
                   mo_DOComprobantePago.exoneraciones = CCur(txtExonerado.Text)
                   mo_DOComprobantePago.fechaEmision = CDate(txtFechaBoleta.Text)    'solo u72
                   'grabar Boleta
                   If lnImporte >= 0 Then
                        If oRsPaciente.State = 1 Then oRsPaciente.Close
                        oRsPaciente.Open "select * from CajaComprobantesPago where nroserie='" & mo_DOComprobantePago.nroSerie & "' and NroDocumento ='" & mo_DOComprobantePago.nrodocumento & "'", oConexion, adOpenKeyset, adLockOptimistic
                        If oRsPaciente.RecordCount = 0 Then
                            If oDllFactUCGestionCaja.CajaComprobantePagoServicioAgregarHISTORICO(mo_DOComprobantePago, mo_doCajaGestion, _
                                       mo_DoFactOrdenServPagos, oRs, sighEntidades.Usuario, _
                                       mo_DoAtencion, Val(mo_cmbIdPuntoCarga.BoundText), mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, _
                                       lnIdReceta) = False Then
                                       lcError1 = lcError1 & " no se pudo grabar Boleta: " & lcSerie & " -" & lcNumero & Chr(13) & Chr(10)
                            End If
                        End If
                   End If
              Loop
              oConexion.Close
              
              Set oConexion = Nothing
              Set oCommand = Nothing
              Set oRsPaciente = Nothing
              Set oRs = Nothing
              Set oRsCatalogo = Nothing
              
              
              MsgBox "terminó " & lcError1
            End If
       End If
       Exit Sub
ErrExc:
      MsgBox lcError & Chr(13) & Err.Number & " - " & Err.Description
      Exit Sub
      Resume

End Sub

Private Sub grdGestionCaja_ClickCellButton(ByVal Cell As UltraGrid.SSCell)
    If lbTienePermisoReimprimeBoleta = True Then
        If MsgBox("Por favor confirmar, ¿Realmente desea REIMPRIMIR  ?", vbQuestion + vbYesNo, "") = vbNo Then
            Exit Sub
        End If
        Dim oImprimeBoletaContinua As New RptCaja
        Dim lcRuc99 As String
        lcRuc99 = IIf(IsNull(grdGestionCaja.ActiveRow.Cells("ruc").Value), "", grdGestionCaja.ActiveRow.Cells("ruc").Value)
        oImprimeBoletaContinua.ImpresionBoletaEnDosTYPE False, DevuelveRUCyDIRECCIONProveedor(False, lcRuc99), _
                                                        "EFECTIVO:       VUELTO:", _
                                                        grdGestionCaja.ActiveRow.Cells("NroSerie").Value, _
                                                        grdGestionCaja.ActiveRow.Cells("NroDocumento").Value, _
                                                        IIf(Val(wxParametro208) = 1910, True, False), _
                                                        IIf(grdGestionCaja.ActiveRow.Cells("IdTipoOrden").Value = 1, _
                                                               sighEntidades.sghServicio, sighEntidades.sghbien), _
                                                        False, False, True
        Set oImprimeBoletaContinua = Nothing
    End If
    
End Sub

'***************daniel barrantes**************
'***************Muestra en PANTALLA el DETALLE un RECIBO elegido previamente
'***************
Private Sub grdGestionCaja_DblClick()
'  Dim oFactOrdenesServicio As New FactOrdenesServicio
'  Dim oDOFactOrdenServicio As New DOFactOrdenServicio
'  Dim oFactOrdenesBienesInsumo As New FactOrdenesBienesInsumo
'  Dim oDOFactOrdenBienInsumo As New DOFactOrdenBienInsumo
'  Dim lIdComprobantePago As Long
'  ml_Estado = grdGestionCaja.ActiveRow.Cells("idEstadoComprobante").Value
   If lbPuedeVerVistaPrevia = False Then
        Select Case grdGestionCaja.ActiveRow.Cells("IdTipoOrden").Value
          Case 1  'Servicios en CAJA SERVICIO
            ImpresionDelRecibo grdGestionCaja.ActiveRow.Cells("NroSerie").Value, grdGestionCaja.ActiveRow.Cells("NroDocumento").Value, sghServicio, sghPantalla, IIf(grdGestionCaja.ActiveRow.Cells("IdTipoComprobante").Value = 2, True, False)
          Case 2  'Bienes e insumos en CAJA FARMACIA
            ImpresionDelRecibo grdGestionCaja.ActiveRow.Cells("NroSerie").Value, grdGestionCaja.ActiveRow.Cells("NroDocumento").Value, sghbien, sghPantalla, IIf(grdGestionCaja.ActiveRow.Cells("IdTipoComprobante").Value = 2, True, False)
          Case 3  'Bienes e insumos en CAJA SERVICIO
            ImpresionDelRecibo grdGestionCaja.ActiveRow.Cells("NroSerie").Value, grdGestionCaja.ActiveRow.Cells("NroDocumento").Value, sghbien, sghPantalla, IIf(grdGestionCaja.ActiveRow.Cells("IdTipoComprobante").Value = 2, True, False)
        End Select
    Else
        Dim oImprimeBoletaContinua As New RptCaja
        Dim lcRuc99 As String
        wxParametro527 = "S"
        wxParametro288 = "S"
        lcRuc99 = IIf(IsNull(grdGestionCaja.ActiveRow.Cells("ruc").Value), "", grdGestionCaja.ActiveRow.Cells("ruc").Value)
        oImprimeBoletaContinua.ImpresionBoletaEnDosTYPE True, DevuelveRUCyDIRECCIONProveedor(False, lcRuc99), _
                                                        "EFECTIVO:       VUELTO:", _
                                                        grdGestionCaja.ActiveRow.Cells("NroSerie").Value, _
                                                        grdGestionCaja.ActiveRow.Cells("NroDocumento").Value, _
                                                        True, _
                                                        IIf(grdGestionCaja.ActiveRow.Cells("IdTipoOrden").Value = 1, _
                                                               sighEntidades.sghServicio, sighEntidades.sghbien), _
                                                        True, True
        Set oImprimeBoletaContinua = Nothing
        wxParametro527 = lcBuscaParametro.SeleccionaFilaParametro(527)
        wxParametro288 = lcBuscaParametro.SeleccionaFilaParametro(288)
    End If
End Sub

Function DevuelveRUCyDIRECCIONProveedor(lbDesdeBotonAceptar As Boolean, lcRuc As String) As String
    Dim lcDireccion1 As String, lcRuc1 As String
    lcDireccion1 = "": lcRuc1 = ""
    If lbDesdeBotonAceptar = True Then
        If txtRuc.Text <> "" Then
           lcRuc1 = txtRuc.Text
           lcDireccion1 = txtDireccionProv.Text
        End If
    ElseIf lcRuc <> "" Then
          Dim oRsTmp As New ADODB.Recordset
          Set oRsTmp = mo_ReglasFacturacion.ProveedoresSeleccionarPorRUC(lcRuc)
          If oRsTmp.RecordCount > 0 Then
             lcDireccion1 = IIf(IsNull(oRsTmp!Direccion), "", oRsTmp!Direccion) & IIf(IsNull(oRsTmp!Email), "", "   (EMAIL: " & oRsTmp!Email & ")")
             lcRuc1 = lcRuc
          End If
          oRsTmp.Close
          Set oRsTmp = Nothing
    End If
    If lcRuc1 = "" Then
       DevuelveRUCyDIRECCIONProveedor = ""
    Else
       DevuelveRUCyDIRECCIONProveedor = "RUC: " & lcRuc1 & "    DIRECCION: " & lcDireccion1
    End If
End Function

'***************daniel barrantes**************
'***************Impresion del RECIBO despues de GRABAR
'***************
Sub ImpresionDelRecibo(lcNroSerie As String, lcNroDcto As String, lnBienFarmacia As sghTipoProducto, lbImpresionFisica As sghImpresion, lbEsFactura As Boolean)
    Dim lbTieneRUC As Boolean
    lbTieneRUC = IIf(txtRuc.Text <> "", True, False)
    Select Case lbImpresionFisica
    Case sghPantalla
            Dim oRecibo As New RecibosBoleta
            oRecibo.EsAnulado = ml_Estado
            oRecibo.lbTienePermisoReimprimeBoleta = lbTienePermisoReimprimeBoleta
            oRecibo.ImprimirDEBB lcNroSerie, lcNroDcto, lnBienFarmacia, ml_idUsuario
            oRecibo.Show 1
            Set oRecibo = Nothing
    Case sghImpresoraBoletaContinua
            Dim oImprimeBoletaContinua As New RptCaja
            Dim lbImprimeEnPantalla As Boolean
            lbImprimeEnPantalla = False
            If wxParametro534 = "S" Then
               lbImprimeEnPantalla = True
            End If
            
            
            
            If mi_Opcion = sghopcionespago.sghPagarCuentaExistente And lbCargaEstadoDeCuentaFarmacia = False Then
               oImprimeBoletaContinua.ImpresionBoletaEnDosTYPE lbImprimeEnPantalla, DevuelveRUCyDIRECCIONProveedor(True, ""), _
                                      lcVuelto, lcNroSerie, lcNroDcto, _
                                      IIf(Val(wxParametro208) = 1910, True, False), _
                                      lnBienFarmacia, True, , lbTieneRUC
            Else
               oImprimeBoletaContinua.ImpresionBoletaEnDosTYPE lbImprimeEnPantalla, DevuelveRUCyDIRECCIONProveedor(True, ""), _
                                      lcVuelto, lcNroSerie, lcNroDcto, _
                                      IIf(Val(wxParametro208) = 1910, True, False), _
                                      lnBienFarmacia, False, , lbTieneRUC
            End If
            Set oImprimeBoletaContinua = Nothing
    Case sghImpresoraBoletaPorBoleta
        Dim oImprimeBoletaDOS As New RptCaja
        oImprimeBoletaDOS.ImpresionBoletaEnDOS lcNroSerie, lcNroDcto, lnBienFarmacia
        Set oImprimeBoletaDOS = Nothing
    End Select
    'kike 2017
    If lbTieneLicenciaParaNotaCreditoYsunat = True Then
        Dim oExportar As New SIGHProxies.Procesos
        oExportar.ExportarFacturasBoletas "", "", lcNroSerie, lcNroDcto, lbUsaResumenDiarioSunat
        Set oExportar = Nothing
    End If
    '
    If lbTieneRUC = True Then
        Dim oProveedores As New Proveedores
        Dim oDoProveedores As New DoProveedores
        Dim oConexion As New Connection
        oConexion.CommandTimeout = 900
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighEntidades.CadenaConexion
        Set oProveedores.Conexion = oConexion
        oDoProveedores.Email = txtEmailProv.Text
        oDoProveedores.Direccion = txtDireccionProv.Text
        oDoProveedores.idProveedor = txtRuc.Tag
        oDoProveedores.IdUsuarioAuditoria = sighEntidades.Usuario
        oDoProveedores.razonSocial = txtRazonSocial.Text
        oDoProveedores.ruc = txtRuc.Text
        If txtRuc.Tag > 0 Then
            If Not oProveedores.Modificar(oDoProveedores) Then
            End If
        Else
            
            If Not oProveedores.Insertar(oDoProveedores) Then
            End If
        End If
        oConexion.Close
        Set oProveedores = Nothing
        Set oDoProveedores = Nothing
        Set oConexion = Nothing
        If Len(txtEmailProv.Text) > 3 And InStr(txtEmailProv.Text, "@") > 1 Then
            Dim mo_email As New SIGHProxies.Procesos
            mo_email.EnviaEmail lcParametro524, lcParametro523, _
                        wxParametro205, _
                        "c:\boleta.txt", txtEmailProv.Text, "Impresión de Factura N° " & lcNroSerie & "-" & lcNroDcto
            Set mo_email = Nothing
        End If
    End If
End Sub


Private Sub grdGestionCaja_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
End Sub

Private Sub grdGestionCaja_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
        Select Case Val(Row.Cells("IdEstadoComprobante").GetText())
        Case 9   'anulado
            Row.Appearance.ForeColor = vbRed
            'Row.Appearance.Font.Strikethrough = True
        Case 6   'Devuelto
            Row.Appearance.ForeColor = vbGreen
        End Select
End Sub







Private Sub optCobrarOrdenExistente_Click(Value As Integer)
    mi_Opcion = sghopcionespago.sghPagarOrdenExistente
    LimpiarOpciones
    LimpiarFormulario
    fraOpciones.Visible = True
    
    mo_Formulario.HabilitarDeshabilitar cmbOrdenes, True
    
    tabFactProductosPorCuenta.Visible = False
    UcFacturacionContado1.Visible = False
    txtNroHistoria.Enabled = True: lbEsDevolucion = False
    ConfiguracionParaPreVenta
    
    'lblOrden.Caption = "N° Orden"                   'debb-18/02/2011
    UcFacturacionContado1.Visible = False
    ucFacturacionProductos.Visible = True
    
    On Error Resume Next
    cmbOrdenes.SetFocus
End Sub

Sub ConfiguracionParaPreVenta()
End Sub



Private Sub optFarmacia_Click(Value As Integer)
    ml_TipoProducto = sghbien
    ucFacturacionProductos.TipoProducto = sghbien
End Sub




Private Sub optNuevoOrdenPagoConHistoria_Click(Value As Integer)
    cmbIdTipoComprobante_Click
    mi_Opcion = sghopcionespago.sghNuevoPagoConHistoria
    LimpiarOpciones
    mo_cmbIdPuntoCarga.BoundText = 99
    LimpiarFormulario
    
    fraOpciones.Visible = True
    mo_Formulario.HabilitarDeshabilitar cmbIdTipoGenHistoriaClinica, True
    mo_Formulario.HabilitarDeshabilitar txtNroHistoria, True
    mo_Formulario.HabilitarDeshabilitar txtNroCuenta, True
    mo_Formulario.HabilitarDeshabilitar txtDni, True  'debb-18/02/2011
    lblBuscaDNIReniec.Visible = False
    
    tabFactProductosPorCuenta.Visible = False
    lbEsDevolucion = False
    cmbIdTipoFinanciamiento.Visible = True
    ChequeaSiF8F9F11F12estanElegido
    InabilitaTotales
    If lbTienePermisoExonerarPacExterno = True Then
        mo_Formulario.HabilitarDeshabilitar txtExonerado, True
    End If
    On Error Resume Next
    txtNroCuenta.SetFocus
End Sub



Sub InabilitaTotales()
    mo_Formulario.HabilitarDeshabilitar txtIngresado, False
    mo_Formulario.HabilitarDeshabilitar txtPendientePago, False
    mo_Formulario.HabilitarDeshabilitar txtExonerado, False
    mo_Formulario.HabilitarDeshabilitar txtPagoACuenta, False
    mo_Formulario.HabilitarDeshabilitar txtTotal, False
    mo_Formulario.HabilitarDeshabilitar txtVuelto, False
    mo_Formulario.HabilitarDeshabilitar txtFalta, False
    mo_Formulario.HabilitarDeshabilitar txtCtaFarmTfarmacia, False
    mo_Formulario.HabilitarDeshabilitar txtCtaFarmExonerado, False
    mo_Formulario.HabilitarDeshabilitar txtCtaServTservicio, False
    mo_Formulario.HabilitarDeshabilitar txtCtaServExonerado, False

End Sub

Private Sub optNuevoOrdenPagoSinHistoria_Click(Value As Integer)
    cmbIdTipoComprobante_Click
    mi_Opcion = sghopcionespago.sghNuevoPagoSinHistoria
    
    fraOpciones.Visible = True
    
    LimpiarOpciones
    
    mo_cmbIdPuntoCarga.BoundText = 99
    
    LimpiarFormulario
    
    mo_Formulario.HabilitarDeshabilitar cmbIdTipoFinanciamiento, True
    mo_Formulario.HabilitarDeshabilitar txtObservaciones, True
    mo_Formulario.HabilitarDeshabilitar txtDni, True     'debb-05/12/2017
    
    lblBuscaDNIReniec.Visible = IIf(wxParametro387 = "S", True, False)
    mo_Formulario.HabilitarDeshabilitar txtDni, True
    
    'frmPreventaServ.Visible = True
    tabFactProductosPorCuenta.Visible = False
    lbEsDevolucion = False
    cmdPaquetes.Enabled = True
    mo_cmbIdTipoFinanciamiento.BoundText = lcBuscaParametro.SeleccionaFilaParametro(260)
    ChequeaSiF8F9F11F12estanElegido

    InabilitaTotales
    If lbTienePermisoExonerarPacExterno = True Then
        mo_Formulario.HabilitarDeshabilitar txtExonerado, True
    End If
    On Error Resume Next
    txtRazonSocial.SetFocus
End Sub




'debb-16/02/2011
Private Sub optOrdenExistenteFS_Click(Value As Integer)
    cmbIdTipoComprobante_Click
    mi_Opcion = sghopcionespago.sghPagarOrdenExistenteFS
    LimpiarOpciones
    LimpiarFormulario
    fraOpciones.Visible = True
    
    mo_Formulario.HabilitarDeshabilitar cmbOrdenes, True
    mo_Formulario.HabilitarDeshabilitar txtDni, True     'debb-05/12/2017
    
    tabFactProductosPorCuenta.Visible = False
    UcFacturacionContado1.Visible = False
    lbEsDevolucion = False
    ConfiguracionParaPreVenta
    
    UcFacturacionContado1.Visible = False
    ucFacturacionProductos.Visible = True
    ChequeaSiF8F9F11F12estanElegido
    
    On Error Resume Next
    cmbOrdenes.SetFocus

End Sub

Private Sub OptOrdenFarmacia_Click(Value As Integer)
    mi_Opcion = sghopcionespago.sghPagarOrdenExistenteF
    LimpiarOpciones
    LimpiarFormulario
    
    fraOpciones.Visible = True
    
    mo_Formulario.HabilitarDeshabilitar cmbOrdenes, True
    mo_Formulario.HabilitarDeshabilitar txtDni, True
    
    tabFactProductosPorCuenta.Visible = False
    UcFacturacionContado1.Visible = False
    txtNroHistoria.Enabled = True: lbEsDevolucion = False
    ConfiguracionParaPreVenta
    
    UcFacturacionContado1.Visible = True
    
    ucFacturacionProductos.Visible = False
    
    On Error Resume Next
    cmbOrdenes.SetFocus

End Sub

'debb-17/02/2011
Private Sub optPagarCtaTotal_Click(Value As Integer)
    cmbIdTipoComprobante_Click
    '
    If lbTienePermisoSoloParaBoletaServicio = True And lbTienePermisoSoloParaBoletaFarmacia = False Then
       optPagarEstadoDeCuenta_Click 1
       Exit Sub
    ElseIf lbTienePermisoSoloParaBoletaFarmacia = True And lbTienePermisoSoloParaBoletaServicio = False Then
       optPagarEstadoDeCTAFarmacia_Click 1
       Exit Sub
    End If
    '
    mi_Opcion = sghopcionespago.sghPagarCuentaTotalFS
    lbCargaEstadoDeCuentaFarmacia = False
    lbCargaEstadoDeCuentaFS = True
    LimpiarOpciones
    LimpiarFormulario
    
    fraOpciones.Visible = False
    
    mo_Formulario.HabilitarDeshabilitar cmbIdTipoGenHistoriaClinica, True
    mo_Formulario.HabilitarDeshabilitar txtNroHistoria, True
    mo_Formulario.HabilitarDeshabilitar txtNroCuenta, True
    mo_Formulario.HabilitarDeshabilitar txtPagoACuenta, True
    mo_Formulario.HabilitarDeshabilitar txtDni, True  'debb-18/02/2011
    lblBuscaDNIReniec.Visible = False
    
    tabFactProductosPorCuenta.Visible = True
    tabFactProductosPorCuenta.TabVisible(0) = True
    tabFactProductosPorCuenta.TabVisible(1) = True
    lbEsDevolucion = False
    ChequeaSiF8F9F11F12estanElegido
    '
    lbCargaEstadoDeCuentaFS = True
    '
    InabilitaTotales
    On Error Resume Next   'debb-10/08/2016
    txtNroCuenta.SetFocus

End Sub

Private Sub optPagarEstadoDeCTAFarmacia_Click(Value As Integer)
    
    mi_Opcion = sghopcionespago.sghPagarCuentaExistente
    lbCargaEstadoDeCuentaFarmacia = True
    LimpiarOpciones
    LimpiarFormulario
    
    fraOpciones.Visible = False
    
    mo_Formulario.HabilitarDeshabilitar cmbIdTipoGenHistoriaClinica, True
    mo_Formulario.HabilitarDeshabilitar txtNroHistoria, True
    mo_Formulario.HabilitarDeshabilitar txtNroCuenta, True
    mo_Formulario.HabilitarDeshabilitar txtPagoACuenta, True
    mo_Formulario.HabilitarDeshabilitar txtDni, True
    
    tabFactProductosPorCuenta.Visible = True
    tabFactProductosPorCuenta.TabVisible(0) = False
    tabFactProductosPorCuenta.TabVisible(1) = True
    txtNroHistoria.Enabled = True: lbEsDevolucion = False
    InabilitaTotales
    txtNroCuenta.SetFocus
End Sub

Private Sub optPagarEstadoDeCuenta_Click(Value As Integer)
    
    mi_Opcion = sghopcionespago.sghPagarCuentaExistente
    lbCargaEstadoDeCuentaFarmacia = False
    LimpiarOpciones
    LimpiarFormulario
    
    fraOpciones.Visible = False
    
    mo_Formulario.HabilitarDeshabilitar cmbIdTipoGenHistoriaClinica, True
    mo_Formulario.HabilitarDeshabilitar txtNroHistoria, True
    mo_Formulario.HabilitarDeshabilitar txtNroCuenta, True
    mo_Formulario.HabilitarDeshabilitar txtPagoACuenta, True
    mo_Formulario.HabilitarDeshabilitar txtDni, True
    
    tabFactProductosPorCuenta.Visible = True
    If optServicios.Value Then
        tabFactProductosPorCuenta.TabVisible(0) = True
        tabFactProductosPorCuenta.TabVisible(1) = False
    Else
        tabFactProductosPorCuenta.TabVisible(0) = False
        tabFactProductosPorCuenta.TabVisible(1) = True
    End If
    txtNroHistoria.Enabled = True: lbEsDevolucion = False
    InabilitaTotales
    txtNroCuenta.SetFocus
End Sub

Private Sub optRealizarAnulacion_Click(Value As Integer)
    mi_Opcion = sghopcionespago.sghAnulacion
    
    fraOpciones.Visible = True
    
    LimpiarOpciones
    LimpiarFormulario
    mo_Formulario.HabilitarDeshabilitar txtNserieB, True
    mo_Formulario.HabilitarDeshabilitar txtNdocumentoB, True
    ChequeaSiF8F9F11F12estanElegido
    
    tabFactProductosPorCuenta.Visible = False
    lbEsDevolucion = False
    txtNreceta.Enabled = False: cmbBuscaReceta.Enabled = False
    
    InabilitaTotales
    txtNserieB.SetFocus
End Sub









Private Sub optRealizarDevolucion_Click(Value As Integer)
    
    lbEsDevolucion = True
    
    mi_Opcion = sghopcionespago.sghDevolucion
    fraOpciones.Visible = True
    LimpiarOpciones
    ucFacturacionProductos.LimpiarGrilla
    mo_Formulario.HabilitarDeshabilitar cmbIdTipoGenHistoriaClinica, False
    mo_Formulario.HabilitarDeshabilitar txtNroHistoria, False
    mo_Formulario.HabilitarDeshabilitar txtNombres, False
    mo_Formulario.HabilitarDeshabilitar cmbOrdenes, True
    mo_Formulario.HabilitarDeshabilitar cmbIdPuntoDeCarga, False
    mo_Formulario.HabilitarDeshabilitar cmbFechaIngreso, False
    mo_Formulario.HabilitarDeshabilitar txtNroCuenta, False
    mo_Formulario.HabilitarDeshabilitar txtNserieB, False
    mo_Formulario.HabilitarDeshabilitar txtNdocumentoB, False

    ucFacturacionProductos.Visible = True
    tabFactProductosPorCuenta.Visible = False
    txtNroHistoria.Enabled = False
End Sub

Private Sub optReimprimirComprobante_Click(Value As Integer)
    
    mi_Opcion = sghopcionespago.sghReimprimirComprobante
    
    fraOpciones.Visible = True
    
    LimpiarOpciones
    ucFacturacionProductos.LimpiarGrilla
    mo_Formulario.HabilitarDeshabilitar cmbIdTipoGenHistoriaClinica, False
    mo_Formulario.HabilitarDeshabilitar txtNroHistoria, False
    mo_Formulario.HabilitarDeshabilitar txtNombres, False
    mo_Formulario.HabilitarDeshabilitar cmbOrdenes, True
    mo_Formulario.HabilitarDeshabilitar cmbIdPuntoDeCarga, False
    mo_Formulario.HabilitarDeshabilitar cmbFechaIngreso, False
    mo_Formulario.HabilitarDeshabilitar txtNroCuenta, False

    ucFacturacionProductos.Visible = True
    tabFactProductosPorCuenta.Visible = False
    txtNroHistoria.Enabled = False: lbEsDevolucion = False
    InabilitaTotales
End Sub

Private Sub optServicios_Click(Value As Integer)
    ml_TipoProducto = sghServicio
    ucFacturacionProductos.TipoProducto = sghServicio
End Sub

Private Sub optSHecogGeneral_Click(Value As Integer)
    If optSHecogGeneral.Value = True Then
       ChequeaSiF3oF7estanElegidos
       LimpiarFormulario
       ucFacturacionProductos.FiltraCpt = sghCptSoloEcografiaGeneral
       On Error Resume Next
       UserControl.txtRazonSocial.SetFocus
    End If

End Sub

Private Sub optSHecogObst_Click(Value As Integer)
    If optSHecogObst.Value = True Then
       ChequeaSiF3oF7estanElegidos
       LimpiarFormulario
       ucFacturacionProductos.FiltraCpt = sghCptSoloEcografiaObstetrica
       On Error Resume Next
       UserControl.txtRazonSocial.SetFocus
    End If

End Sub

Private Sub optSHLaboratorio_Click(Value As Integer)
    If optSHLaboratorio.Value = True Then
       ChequeaSiF3oF7estanElegidos
       LimpiarFormulario
       ucFacturacionProductos.FiltraCpt = sghCptSoloLaboratorio
       On Error Resume Next
       UserControl.txtRazonSocial.SetFocus
    End If
End Sub

Sub ChequeaSiF3oF7estanElegidos()
       If optNuevoOrdenPagoConHistoria.Value = False Then
          If optNuevoOrdenPagoSinHistoria.Value = False Then
              optNuevoOrdenPagoSinHistoria.Value = True
              optNuevoOrdenPagoSinHistoria_Click True
          End If
       End If
End Sub

Sub ChequeaSiF8F9F11F12estanElegido()
    If optNuevoOrdenPagoConHistoria.Value = True Or optNuevoOrdenPagoSinHistoria.Value = True Then
        If optSHLaboratorio.Value = True Then
           optSHLaboratorio_Click True
        ElseIf optSHrayosX.Value = True Then
           optSHrayosX_Click True
        ElseIf optSHtomografia.Value = True Then
           optSHtomografia_Click True
        ElseIf optSHecogGeneral.Value = True Then
           optSHecogGeneral_Click True
        ElseIf optSHecogObst.Value = True Then
           optSHecogObst_Click True
        Else
           optSHtodos.Value = True
           optSHtodos_Click True
        End If
        'FraServHosp.BackColor = &H8000000F
        mo_Formulario.HabilitarDeshabilitar FraServHosp, True
    Else
        'FraServHosp.BackColor = &HF9EADF
        mo_Formulario.HabilitarDeshabilitar FraServHosp, False
    End If
End Sub


Private Sub optSHrayosX_Click(Value As Integer)
    If optSHrayosX.Value = True Then
       ChequeaSiF3oF7estanElegidos
       LimpiarFormulario
       ucFacturacionProductos.FiltraCpt = sghCptSoloRayosX
       On Error Resume Next
       UserControl.txtRazonSocial.SetFocus
    End If
End Sub



Private Sub optSHtodos_Click(Value As Integer)
    ucFacturacionProductos.FiltraCpt = sghMuestraTodosCpt
    On Error Resume Next
    UserControl.txtRazonSocial.SetFocus
End Sub

Private Sub optSHtomografia_Click(Value As Integer)
    If optSHtomografia.Value = True Then
       ChequeaSiF3oF7estanElegidos
       LimpiarFormulario
       ucFacturacionProductos.FiltraCpt = sghCptSoloTomografia
       On Error Resume Next
       UserControl.txtRazonSocial.SetFocus
    End If

End Sub



Private Sub tabFactProductosPorCuenta_KeyUp(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

'debb-16/02/2011
Private Sub tabGestionCaja_Click(PreviousTab As Integer)
    If PreviousTab = 0 Then
       UserControl.KeyPreview = True
    Else
       UserControl.KeyPreview = False
    End If
End Sub




Private Sub txtDNI_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtDni
  
End Sub

Private Sub txtDNI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txtDni_LostFocus
    End If
End Sub

Private Sub txtDni_KeyUp(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

'debb-18/02/2011
Private Sub txtDni_LostFocus()
    If txtDni.Text <> "" Then
       'debb-05/12/2017 (inicio)
       lcDNIbuscado = txtDni.Text
       If mi_Opcion = sghopcionespago.sghNuevoPagoSinHistoria Or mi_Opcion = sghopcionespago.sghPagarOrdenExistenteF Then
          Dim oRsTmp1 As New Recordset
          Set oRsTmp1 = mo_AdminAdmision.PacientesSeleccionarXdni(lcDNIbuscado, 1)
          If oRsTmp1.RecordCount > 0 Then
             txtRazonSocial.Text = oRsTmp1!ApellidoPaterno & " " & Trim(oRsTmp1!ApellidoMaterno) & " " & Trim(oRsTmp1!PrimerNombre)
             lnIdPacienteDelDNIelegido = oRsTmp1!idPaciente
             txtRazonSocial.Enabled = False
             
          End If
          oRsTmp1.Close
          Set oRsTmp1 = Nothing
          Exit Sub
       End If
       'debb-05/12/2017 (fin)



       txtNroCuenta.Text = ""
       LimpiarFormulario
       If Len(txtDni.Text) = 8 Then
          Dim oRsTmp As New Recordset
          Dim lnIdPacienteHallado As Long
          Dim oConexion As New Connection
          oConexion.Open sighEntidades.CadenaConexion
          oConexion.CursorLocation = adUseClient
          Set oRsTmp = mo_AdminAdmision.PacientesXdni(txtDni.Text, oConexion)
          If oRsTmp.RecordCount > 0 Then
             lnIdPacienteHallado = oRsTmp.Fields!idPaciente
             If Not IsNull(oRsTmp.Fields!idTipoNumeracion) Then
                mo_cmbIdTipoGenHistoriaClinica.BoundText = oRsTmp.Fields!idTipoNumeracion
             End If

             UserControl.txtNroHistoria.Text = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(oRsTmp.Fields!NroHistoriaClinica)), False)
             txtRazonSocial.Text = Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & " " & Trim(oRsTmp.Fields!PrimerNombre) & " " & Trim(oRsTmp.Fields!SegundoNombre)
             txtRazonSocial.Enabled = False
             oRsTmp.Close
             Set oRsTmp = Nothing
          Else
                If wxParametro387 = "S" And mi_Opcion = sghNuevoPagoSinHistoria Then 'BUSCAR EN RENIEC
                    mo_Reniec.Inicializar
                    mo_Reniec.ConsultarDNIenReniec Trim(txtDni.Text)
                    If mo_Reniec.ApellidoPaterno <> "" Then
                        txtRazonSocial.Text = Trim(mo_Reniec.ApellidoPaterno) & " " & Trim(mo_Reniec.ApellidoMaterno) & " " & Trim(mo_Reniec.PrimerNombre) & " " & Trim(mo_Reniec.SegundoNombre)
                    Else
                        MsgBox "Ese Nro de DNI no existe en RENIEC", vbInformation, "Caja"
                    End If
                    Exit Sub
                Else
                    MsgBox "Ese Nro de DNI no existe", vbInformation, "Caja"
                    oRsTmp.Close
                    Set oRsTmp = Nothing
                End If
          End If
          oConexion.Close
          Set oConexion = Nothing
       Else
          MsgBox "El DNI debe tener 8 digitos", vbInformation, "Caja"
       End If
    End If
End Sub

Private Sub txtEfectivo_Change()
Dim dDiferencia As Double

    If txtEfectivo = "" Or txtEfectivo = "," Or CCur(txtTotal.Text) = 0 Then
        Exit Sub
    End If
    md_Total = CCur(txtTotal.Text)
    dDiferencia = CDbl(txtEfectivo.Text) - md_Total
    
    If dDiferencia > 0 Then
        txtVuelto = IIf(dDiferencia = 0, "", Format(dDiferencia, "#######.#0"))
        txtFalta = ""
    Else
        txtVuelto = ""
        txtFalta = IIf(dDiferencia = 0, "", Format(dDiferencia, "#######.#0"))
    End If

End Sub

Private Sub txtEfectivo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       btnAceptar.SetFocus
       Exit Sub
    End If
    If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
           KeyAscii = 0
       End If
    End If

End Sub







Private Sub txtEfectivo_KeyUp(KeyCode As Integer, Shift As Integer)
   AdministrarKeyPreview KeyCode
End Sub

Private Sub txtExonerado_Change()
    ActualizaTotalApagar
End Sub

Private Sub txtExonerado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txtExonerado_LostFocus
    End If
End Sub

Private Sub txtExonerado_KeyUp(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtExonerado_LostFocus()
    If Val(txtExonerado.Text) > 0 And (mi_Opcion = sghopcionespago.sghNuevoPagoSinHistoria Or _
                                        mi_Opcion = sghopcionespago.sghNuevoPagoConHistoria Or _
                                            mi_Opcion = sghopcionespago.sghPagarOrdenExistenteF) Then     'debb-04/09/2018
       cmbServicioSocial.Visible = True: txtServicioSocial.Visible = True: lblServicioSocial.Visible = True
       mo_cmbServicioSocial.BoundText = ""
       txtServicioSocial.SetFocus
    Else
       cmbServicioSocial.Visible = False: txtServicioSocial.Visible = False: lblServicioSocial.Visible = False
    End If
End Sub



Private Sub txtFalta_KeyUp(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtFdesde_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFdesde

End Sub



Private Sub txtFdesde_LostFocus()
    If Not IsDate(txtFdesde.Text) Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        'txtFdesde.Text = SIGHEntidades.FECHA_VACIA_DMY_HM
        btnLimpiar_Click
        Exit Sub
    End If

End Sub

Private Sub txtFhasta_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFhasta

End Sub

Private Sub txtFhasta_LostFocus()
    If Not IsDate(txtFhasta.Text) Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        'txtFhasta.Text = SIGHEntidades.FECHA_VACIA_DMY_HM
        btnLimpiar_Click
        Exit Sub
    End If

End Sub

Private Sub txtNdocumentoB_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtNdocumentoB_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Val(txtNdocumentoB.Text) > 0 Then
       cmdLeer_Click
    End If
End Sub



Private Sub txtNdocumentoB_KeyUp(KeyCode As Integer, Shift As Integer)
     AdministrarKeyPreview KeyCode

End Sub



Private Sub txtNreceta_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           txtNreceta_LostFocus
        End If
End Sub

Private Sub txtNreceta_KeyUp(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNreceta_LostFocus()
    If Val(txtNreceta.Text) > 0 Then
       lbEsUnaRecetaOtrosCpt = False
       If Len(txtNreceta.Text) > 9 Then
          MsgBox "La Receta no puede pasar de 9 nùmeros", vbInformation, "Caja"
          txtNreceta.Text = ""
          Exit Sub
       End If
       On Error GoTo ErrReceta
       Dim lcSql As String
       Dim oRsTmp1 As New Recordset, lnRecetaProcesada As Long, lnCuenta As Long
       lnRecetaProcesada = Val(txtNreceta.Text)
       '
        Set oRsTmp1 = mo_ReglasComunes.RecetasConCabeceraYdetalleSoloCpt(lnRecetaProcesada, sghRecetaEstados.sighRecetaRegistrada)
        If oRsTmp1.RecordCount > 0 Then
          lnCuenta = oRsTmp1.Fields!idCuentaAtencion
          If oRsTmp1.Fields!idPuntoCarga = sghPtoCargaFarmacia Then
                MsgBox "Es una receta de FARMACIA, tiene que dirigirse a Farmacia", vbInformation, "Caja"
          Else
                If oRsTmp1.Fields!idEstado <> sghRecetaEstados.sighRecetaRegistrada Then
                    mo_ReglasComunes.RecetaChequeaEstadoActual oRsTmp1.Fields!idCuentaAtencion, _
                                                           oRsTmp1.Fields!idEstado, _
                                                           0, oRsTmp1.Fields!DocumentoDespacho

                Else
                    UserControl.optNuevoOrdenPagoConHistoria.Value = True
                    optNuevoOrdenPagoConHistoria_Click True
                    txtNroCuenta.Text = oRsTmp1.Fields!idCuentaAtencion
                    txtNroCuenta_KeyPress 13
                    Select Case oRsTmp1.Fields!idPuntoCarga
                    Case sghPtoCargaRayosX
                         optSHrayosX.Value = True
                         optSHrayosX_Click True
                    Case sghPtoCargaEcogObstetrica
                         optSHecogObst.Value = True
                         optSHecogObst_Click True
                    Case sghPtoCargaEcogGeneral
                         optSHecogGeneral.Value = True
                         optSHecogGeneral_Click True
                    Case sghPtoCargaTomografia
                         optSHtomografia.Value = True
                         optSHtomografia_Click True
                    Case sghPtoCargaServicioHospitalizacion
                         lbEsUnaRecetaOtrosCpt = True
                    Case Else   'Laboratorio
                         optSHLaboratorio.Value = True
                         optSHLaboratorio_Click True
                    End Select
                    
                    mo_cmbIdTipoComprobante.BoundText = wxIdTipoComprobanteDefault 'FrankSunat
                    'cmbIdTipoComprobante_Click 'FrankSunat
                    If CuentaTieneSeguro(oRsTmp1.Fields!idCuentaAtencion) = False Then '---FRANK 11/11/2015
                        ucFacturacionProductos.CargaProductosPorIdReceta oRsTmp1
                        lnIdReceta = lnRecetaProcesada '---FRANK 11/11/2015
                    End If

                End If
          End If
       Else
          MsgBox "Ese N° Receta NO EXISTE", vbInformation, "Caja"
       End If
       oRsTmp1.Close
       Set oRsTmp1 = Nothing
    End If
    Exit Sub
ErrReceta:
    MsgBox Err.Description
    Resume
End Sub

Public Function CuentaTieneSeguro(lnIdCuentaAtencion As Long) As Boolean
    Dim oConexion As New Connection
    Dim oRsTmp As New Recordset
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    CuentaTieneSeguro = False
    Set oRsTmp = mo_AdminAdmision.atencionesXtipoFinanciamiento(lnIdCuentaAtencion, oConexion)
    If oRsTmp.RecordCount > 0 Then
       mo_cmbIdTipoFinanciamiento.BoundText = oRsTmp.Fields!IdFormaPago
       lblCuentaConSeguro.Caption = IIf(oRsTmp.Fields!generaPago = 1, "", "Con Seguro")
       If oRsTmp.Fields!generaPago <> 1 Then CuentaTieneSeguro = True
    End If
    oRsTmp.Close
    
    Set oRsTmp = Nothing
    oConexion.Close
    Set oConexion = Nothing
End Function

'Private Sub txtNroCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
'   mo_Teclado.RealizarNavegacion KeyCode, txtNroCuenta
'End Sub

Private Sub txtNroCuenta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      
      If mo_Teclado.TextoEsSoloNumeros(txtNroCuenta.Text) Then
            If Len(txtNroCuenta.Text) > 9 Then
               MsgBox "El Nro Cuenta no debe pasar de 9 nùmeros", vbInformation, "caja"

               txtNroCuenta.Text = ""
               Exit Sub
            End If
            If optNuevoOrdenPagoConHistoria.Value = True Or optPagarCtaTotal.Value = True Then
                If ChequeaQueCuentaEsPagoSoloDeCita(Val(txtNroCuenta.Text)) = True Then
                   Exit Sub
                End If
            End If
            If optPagarEstadoDeCuenta.Value = True Or optPagarEstadoDeCTAFarmacia.Value = True Or optPagarCtaTotal.Value = True Then
               ml_idCuentaAtencion = Val(txtNroCuenta.Text)
               cmdLeer_Click
            ElseIf optNuevoOrdenPagoConHistoria.Value = True Then
               BuscaPagoConHistoriaUsandoNroCuenta
               If txtNroHistoria.Text <> "" Then
                  ucFacturacionProductos.idCuentaAtencion = Val(txtNroCuenta.Text)
               End If
               'ucFacturacionProductos.TabEnDescripcion
            End If
            MuestraSiTieneSeguro
      End If
   'Else
   End If
End Sub

Sub BuscaPagoConHistoriaUsandoNroCuenta()
    Dim oPaciente As New doPaciente
    Dim rsRespuesta As New Recordset
    Dim lcSql As String

    ml_idCuentaAtencion = Val(txtNroCuenta.Text)
    Set rsRespuesta = mo_ReglasFarmacia.FacturacionCuentasAtencionXidCuenta(ml_idCuentaAtencion)
    If rsRespuesta.RecordCount = 0 Then
       Exit Sub
    End If
    lcDNIbuscado = IIf(IsNull(rsRespuesta!nrodocumento), "", rsRespuesta!nrodocumento)
    oPaciente.NroHistoriaClinica = rsRespuesta.Fields!NroHistoriaClinica
    oPaciente.idTipoNumeracion = rsRespuesta.Fields!idTipoNumeracion
    txtNroHistoria.Text = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rsRespuesta.Fields!NroHistoriaClinica)), False)
    
    txtNombres = rsRespuesta!ApellidoPaterno + " " + rsRespuesta!ApellidoMaterno + " " + rsRespuesta!PrimerNombre
    txtRazonSocial.Text = txtNombres
    ml_IdPaciente = rsRespuesta!idPaciente
    rsRespuesta.Close
    'Carga al objeto mo_DoAtencion valores de la ultima atención registrada en la tabla Atenciones de la BD
    Set mo_DoAtencion = mo_ReglasFacturacion.SeleccionarUltimaAtencion(ml_IdPaciente, ml_idCuentaAtencion)
    mo_cmbIdTipoFinanciamiento.BoundText = mo_DoAtencion.IdFormaPago
    If mo_cmbIdTipoFinanciamiento.BoundText = "" Then
       mo_cmbIdTipoFinanciamiento.BoundText = "1"
    End If
    Set ucFacturacionProductos.Atencion = mo_DoAtencion
    
    Set oPaciente = Nothing
    Set rsRespuesta = Nothing
    
End Sub







Private Sub txtNroCuenta_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not (KeyCode = vbKeyF3 Or KeyCode = vbKeyF6) Then
        AdministrarKeyPreview KeyCode
    End If
End Sub

Private Sub txtNroCuenta_LostFocus()
        If optNuevoOrdenPagoConHistoria.Value = True Then
            ucFacturacionProductos.TabEnDescripcion
        End If
        MuestraSiTieneSeguro
End Sub

Sub MuestraSiTieneSeguro()
    lblCuentaConSeguro.Caption = ""
    If Val(txtNroCuenta.Text) > 0 Then
        Dim oConexion As New Connection
        Dim oRsTmp As New Recordset
        oConexion.CommandTimeout = 900
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighEntidades.CadenaConexion
        Set oRsTmp = mo_AdminAdmision.atencionesXtipoFinanciamiento(Val(txtNroCuenta.Text), oConexion)
        If oRsTmp.RecordCount > 0 Then
           lblCuentaConSeguro.Caption = IIf(oRsTmp.Fields!generaPago = 1, "", "Con Seguro")
        End If
        oRsTmp.Close
        oConexion.Close
        Set oRsTmp = Nothing
        Set oConexion = Nothing
     End If
End Sub

Private Sub txtNroDocumentoBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNroDocumentoBusqueda
   AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNroHistoria_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNroHistoria
   AdministrarKeyPreview KeyCode

End Sub


Private Sub txtNroHistoria_KeyUp(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNroHistoriaBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNroHistoriaBusqueda
   AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNroHistoriaBusqueda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       btnBuscar_Click
    End If
End Sub

Private Sub txtNroSerieBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNroSerieBusqueda
   AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNserieB_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNserieB
End Sub

Private Sub txtNserieB_KeyPress(KeyAscii As Integer)
'       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
'           KeyAscii = 0
'       End If

End Sub

Private Sub txtNserieB_KeyUp(KeyCode As Integer, Shift As Integer)
     AdministrarKeyPreview KeyCode
End Sub

Private Sub txtPagoACuenta_Change()
   ActualizaTotalApagar
   md_Total = txtTotal.Text
End Sub


Private Sub txtPagoACuenta_KeyDown(KeyCode As Integer, Shift As Integer)
         AdministrarKeyPreview KeyCode
End Sub


Private Sub txtPagoACuenta_KeyUp(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtRazonSocial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If optCobrarOrdenExistente.Value = True Then
          btnAceptar.SetFocus
          'debb-16/02/2011
       ElseIf optCobrarOrdenExistente.Value = True Or OptOrdenFarmacia.Value = True Or optOrdenExistenteFS.Value = True Then
          btnAceptar.SetFocus
       Else
          ucFacturacionProductos.TabEnDescripcion
       End If
    End If
End Sub


Private Sub txtRazonSocial_KeyUp(KeyCode As Integer, Shift As Integer)
     AdministrarKeyPreview KeyCode

End Sub

Private Sub TxtRsocial_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, TxtRsocial
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtRuc_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtRuc
End Sub

Private Sub txtRuc_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       ElseIf KeyAscii = 13 Then
           txtRuc_LostFocus
       End If
End Sub

Private Sub txtRuc_KeyUp(KeyCode As Integer, Shift As Integer)
     AdministrarKeyPreview KeyCode

End Sub



Private Sub txtRuc_LostFocus()
    If txtRuc.Text <> "" Then
          Dim oRsTmp As New ADODB.Recordset
          Set oRsTmp = mo_ReglasFacturacion.ProveedoresSeleccionarPorRUC(txtRuc.Text)
          If oRsTmp.RecordCount > 0 Then
             txtRazonSocial.Text = oRsTmp.Fields!razonSocial
             txtEmailProv.Text = IIf(IsNull(oRsTmp.Fields!Email), "", oRsTmp.Fields!Email)
             txtRuc.Tag = oRsTmp!idProveedor
             txtDireccionProv.Text = IIf(IsNull(oRsTmp!Direccion), "", oRsTmp!Direccion)
          Else
             txtRuc.Tag = 0
          End If
          oRsTmp.Close
          Set oRsTmp = Nothing
    End If
End Sub

Private Sub txtServicioSocial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txtServicioSocial_LostFocus
    End If
End Sub

Private Sub txtServicioSocial_KeyUp(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtServicioSocial_LostFocus()
    If Val(txtServicioSocial.Text) > 0 Then
       mo_cmbServicioSocial.BoundText = txtServicioSocial.Text
       btnAceptar.SetFocus
    End If
End Sub


Private Sub txtVuelto_KeyUp(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub ucFactBienesPorCuenta_SePresionoTeclaEspecial(KeyCode As Integer)
     If KeyCode >= vbKeyF2 And KeyCode <= vbKeyF12 Then
        AdministrarKeyPreview KeyCode
     End If

End Sub

Private Sub ucFactBienesPorCuenta_Totalizado(TotalIngresado As Double, TotalPendientePago As Double, TotalPagoACuenta As Double, TotalExonerado As Double, dTotalPagado As Double, dTotalPorDevolver As Double, dTotalDevuelto As Double, dTotalAnulado As Double)
    
    Select Case mi_Opcion
    Case sghopcionespago.sghPagarCuentaExistente
        lnTotalGrid = TotalIngresado
        md_Ingresado = md_Ingresado + TotalIngresado
        md_PendientePago = md_PendientePago + TotalPendientePago
        md_PagoACuenta = md_PagoACuenta + TotalPagoACuenta
        md_Exonerado = md_Exonerado + TotalExonerado
        md_Total = md_Total + TotalIngresado + TotalPendientePago - TotalPagoACuenta - TotalExonerado
    
        txtIngresado = IIf(md_Ingresado = 0, "", Format(md_Ingresado, "#######.#0"))
        txtPendientePago = IIf(md_PendientePago = 0, "", Format(md_PendientePago, "#######.#0"))
        txtPagoACuenta = IIf(md_PagoACuenta = 0, "0", Format(md_PagoACuenta, "#######.#0"))
        txtExonerado = IIf(md_Exonerado = 0, "0", Format(md_Exonerado, "#######.#0"))
        txtTotal.Text = IIf(md_Total = 0, "0", Format(md_Total, "#######.#0"))
    
    End Select
    ActualizaTotalApagar
End Sub

Private Sub ucFactServiciosPorCuenta_SePresionoTeclaEspecial(KeyCode As Integer)
     If KeyCode >= vbKeyF2 And KeyCode <= vbKeyF12 Then
        AdministrarKeyPreview KeyCode
     End If

End Sub

Private Sub ucFactServiciosPorCuenta_Totalizado(TotalIngresado As Double, TotalPendientePago As Double, TotalPagoACuenta As Double, TotalExonerado As Double, dTotalPagado As Double, dTotalPorDevolver As Double, dTotalDevuelto As Double, dTotalAnulado As Double)

    Select Case mi_Opcion
    Case sghopcionespago.sghPagarCuentaExistente
        lnTotalGrid = TotalIngresado
        md_Ingresado = TotalIngresado
        md_PendientePago = TotalPendientePago
        md_PagoACuenta = TotalPagoACuenta
        md_Exonerado = TotalExonerado
        md_Total = TotalIngresado + TotalPendientePago - TotalPagoACuenta - TotalExonerado
        
        'En caso no hay bienes esto llena los textbox
        txtIngresado = IIf(md_Ingresado = 0, "", Format(md_Ingresado, "#######.#0"))
        txtPendientePago = IIf(md_PendientePago = 0, "", Format(md_PendientePago, "#######.#0"))
        txtPagoACuenta = IIf(md_PagoACuenta = 0, "0", Format(md_PagoACuenta, "#######.#0"))
        txtExonerado = IIf(md_Exonerado = 0, "0", Format(md_Exonerado, "#######.#0"))
        txtTotal.Text = IIf(md_Total = 0, "0", Format(md_Total, "#######.#0"))
    
    End Select
    ActualizaTotalApagar
End Sub





Private Sub UcFacturacionContado1_SePresionoTeclaEspecial(KeyCode As Integer)
     If KeyCode >= vbKeyF2 And KeyCode <= vbKeyF12 Then
        AdministrarKeyPreview KeyCode
     End If

End Sub

Private Sub ucFacturacionProductos_SePresionoTeclaEspecial(KeyCode As Integer)
     If KeyCode >= vbKeyF2 And KeyCode <= vbKeyF12 Then
        AdministrarKeyPreview KeyCode
     End If
End Sub

Private Sub ucFacturacionProductos_Totalizado(TotalIngresado As Double, TotalPendientePago As Double, TotalPagoACuenta As Double, TotalExonerado As Double, dTotalPagado As Double, dTotalPorDevolver As Double, dTotalDevuelto As Double, dTotalAnulado As Double)
    
    Select Case mi_Opcion
    Case sghopcionespago.sghNuevoPagoConHistoria, sghopcionespago.sghNuevoPagoSinHistoria, sghopcionespago.sghPagarOrdenExistente, sghopcionespago.sghPagarCuentaExistente
        lnTotalGrid = TotalIngresado
    Case sghopcionespago.sghAnulacion
        lnTotalGrid = dTotalPagado
        md_Total = dTotalPagado
        txtTotal.Text = IIf(md_Total = 0, "", Format(md_Total, "#######.#0"))
        txtVuelto.Text = txtTotal.Text
    Case sghopcionespago.sghDevolucion
        lnTotalGrid = dTotalPorDevolver
        md_Total = dTotalPorDevolver
        txtTotal.Text = IIf(md_Total = 0, "", Format(md_Total, "#######.#0"))
        txtVuelto.Text = txtTotal.Text
    End Select
    ActualizaTotalApagar
End Sub

Private Sub ucGestionCajaFact1_SeIngresoDescripcion(lcTexto As String)
    UserControl.txtObservaciones.Text = lcTexto
End Sub

Private Sub ucGestionCajaFact1_SeIngresoImporte(lnImporte As Double, lnMontoIGV As Double, lbEsCredito As Boolean)
      ucFacturacionProductos.CajaConDescripcionLargaActualizaImporte lnImporte
      lnMontoIGV99 = lnMontoIGV
      lbTieneCredito99 = lbEsCredito
End Sub





Private Sub UserControl_GotFocus()
   On Error Resume Next
   If lbSeAperturoCAJA = True Then
      lbSeAperturoCAJA = False
      optRealizarAnulacion.Value = True
      optRealizarAnulacion_Click 1
   End If
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
   
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

'debb-16/02/2011
Sub AdministrarKeyPreview(KeyCode As Integer)
    
    Select Case KeyCode
    Case vbKeyEscape
    Case vbKeyF1
          If UserControl.tabGestionCaja.TabVisible(2) = True Then tabGestionCaja.Tab = 2 'Frank 24082015
    Case vbKeyF2
         'Frank 24082015 ''''''''''''''
         If tabGestionCaja.Tab = 2 Then
            btnAceptarPagoNota_Click
         Else
            btnAceptar_Click
         End If
         '''''''''''''''''''''''''''''''
    Case vbKeyF3
          If tabGestionCaja.TabVisible(1) = True Then tabGestionCaja.Tab = 1 'Frank 24082015
          If optNuevoOrdenPagoConHistoria.Enabled = True Then
                optNuevoOrdenPagoConHistoria.Value = True
                optNuevoOrdenPagoConHistoria_Click True
          End If
     Case vbKeyF4
          If tabGestionCaja.TabVisible(1) = True Then tabGestionCaja.Tab = 1 'Frank 24082015
          If optOrdenExistenteFS.Enabled = True Then
                optOrdenExistenteFS.Value = True
                optOrdenExistenteFS_Click True
          End If
     Case vbKeyF5
          If tabGestionCaja.TabVisible(1) = True Then tabGestionCaja.Tab = 1 'Frank 24082015
          If optRealizarAnulacion.Enabled = True Then
                optRealizarAnulacion.Value = True
                optRealizarAnulacion_Click True
          End If
     Case vbKeyF6
          If tabGestionCaja.TabVisible(1) = True Then tabGestionCaja.Tab = 1 'Frank 24082015
          If optPagarCtaTotal.Enabled = True Then
                optPagarCtaTotal.Value = True
                optPagarCtaTotal_Click True
          End If
     Case vbKeyF7
           If tabGestionCaja.TabVisible(1) = True Then tabGestionCaja.Tab = 1 'Frank 24082015
           If optNuevoOrdenPagoSinHistoria.Enabled = True Then
                optNuevoOrdenPagoSinHistoria.Value = True
                optNuevoOrdenPagoSinHistoria_Click True
           End If
     Case vbKeyF8
           If tabGestionCaja.TabVisible(1) = True Then tabGestionCaja.Tab = 1 'Frank 24082015
           optSHLaboratorio.Value = True
           optSHLaboratorio_Click True
'            optNuevoOrdenPagoSinHistoria.Value = True
'            optNuevoOrdenPagoSinHistoria_Click True
'            optSHLaboratorio.Value = True
     Case vbKeyF9
           If tabGestionCaja.TabVisible(1) = True Then tabGestionCaja.Tab = 1 'Frank 24082015
           optSHrayosX.Value = True
           optSHrayosX_Click True
'            optNuevoOrdenPagoSinHistoria.Value = True
'            optNuevoOrdenPagoSinHistoria_Click True
'            optSHrayosX.Value = True
     Case vbKeyF11
          If tabGestionCaja.TabVisible(1) = True Then tabGestionCaja.Tab = 1 'Frank 24082015
          optSHecogGeneral.Value = True
          optSHecogGeneral_Click True
'             optNuevoOrdenPagoSinHistoria.Value = True
'             optNuevoOrdenPagoSinHistoria_Click True
'             optSHecogGeneral.Value = True
     Case vbKeyF12
          If tabGestionCaja.TabVisible(1) = True Then tabGestionCaja.Tab = 1 'Frank 24082015
          optSHecogObst.Value = True
          optSHecogObst_Click True
'            optNuevoOrdenPagoSinHistoria.Value = True
'            optSHecogObst.Value = True
'            optNuevoOrdenPagoSinHistoria_Click True
    End Select
End Sub


Sub ConfigurarGrilla()
    grdGestionCaja.Bands(0).Columns("Turno").Width = 800      '1200
    grdGestionCaja.Bands(0).Columns("Fecha").Width = 1600
    grdGestionCaja.Bands(0).Columns("NroSerie").Width = 1500
    grdGestionCaja.Bands(0).Columns("Turno").Style = ssStyleButton
    
    'FRANK 24082015
    grdNotaCredito.Bands(0).Columns("Turno").Width = 800
    grdNotaCredito.Bands(0).Columns("Fecha").Width = 1600
    grdNotaCredito.Bands(0).Columns("NroSerie").Width = 1500
End Sub

Public Sub ConfigurarCaja()
    
    mo_cmbIdCaja.BoundColumn = "IdCaja"
    mo_cmbIdCaja.ListField = "Descripcion"
    Set mo_cmbIdCaja.RowSource = mo_AdminCaja.CajaSeleccionarTodosParaLista()
    
    mo_cmbIdCajaBusqueda.BoundColumn = "IdCaja"
    mo_cmbIdCajaBusqueda.ListField = "Descripcion"
    Set mo_cmbIdCajaBusqueda.RowSource = mo_AdminCaja.CajaSeleccionarTodosParaLista()
    
End Sub

Public Sub ConfigurarTurno()
    Dim oRsPermisos As New Recordset

    mo_cmbIdTurno.BoundColumn = "IdTurno"
    mo_cmbIdTurno.ListField = "Descripcion"
    Set mo_cmbIdTurno.RowSource = mo_AdminCaja.TurnosSeleccionarTodosParaLista()
    
    mo_cmbIdTurnoBusqueda.BoundColumn = "IdTurno"
    mo_cmbIdTurnoBusqueda.ListField = "Descripcion"
    Set mo_cmbIdTurnoBusqueda.RowSource = mo_AdminCaja.TurnosSeleccionarTodosParaLista()
    
    Set oRsCajeros = mo_AdminCaja.CajerosSeleccionarTodos()
    mo_cmbIdResponsable.BoundColumn = "IdEmpleado"
    mo_cmbIdResponsable.ListField = "DCajero"
    Set mo_cmbIdResponsable.RowSource = oRsCajeros
    If oRsCajeros.RecordCount > 0 Then
    
        Set oRsPermisos = mo_ReglasSeguridad.UsuariosRolesSeleccionarPermisosTodos(Val(sighEntidades.Usuario))
        oRsPermisos.Filter = "idPermiso=1000"
        mo_Formulario.HabilitarDeshabilitar cmbIdResponsable, True
        If oRsPermisos.RecordCount > 0 Then
            mo_cmbIdResponsable.BoundText = sighEntidades.Usuario
            mo_Formulario.HabilitarDeshabilitar cmbIdResponsable, False
        End If
        oRsPermisos.Close
        
'       oRsCajeros.MoveFirst
'       oRsCajeros.Find "idEmpleado=" & sighentidades.Usuario
'       If Not oRsCajeros.EOF Then
'          mo_cmbIdResponsable.BoundText = sighentidades.Usuario
'          mo_Formulario.HabilitarDeshabilitar cmbIdResponsable, False
'       End If
    End If
    Set oRsPermisos = Nothing
End Sub

Public Sub RealizarBusqueda()
    Dim lcFechaIni As Date: Dim lcFechaFin As Date
    Dim lnTotalRecaudado As Double, lnTotBoletas As Double, lnTotFacturas As Double, lnTotDctosAnulados As Double, lnTotDevNotaCred As Double 'Frank 28082015
    Dim lnNroBoletas As Long, lnNroFacturas As Long, lnNroDocumentos As Long, lnNroDctosAnulados As Long, lnNroDevNotaCred As Long 'Frank 28082015
    lcFechaIni = CDate(txtFdesde.Text)
    lcFechaFin = CDate(txtFhasta.Text)
    
    If chkSoloCredito.Value = 0 Then
       Set oRsBusquedaRecibos = mo_AdminCaja.CajaComprobantePagoSeleccionarPorFechaOdocumento("", "", lcFechaIni, lcFechaFin)
    Else
       Set oRsBusquedaRecibos = mo_AdminCaja.CajaComprobantePagoFiltroPorNroSerieDocumentoOporRangoFemision("", "", lcFechaIni, lcFechaFin)
    End If
    
    'Frank 24082015
    Set oRsBusquedaDevNotaCredito = mo_AdminCaja.NotaCreditoDevueltosPorNumYFecha("", "", lcFechaIni, lcFechaFin)
    
    ''''''''''''''''
    ms_MensajeError = ""
    txtTotalCajero.Caption = "0"
    UserControl.lblNroFacturas.Caption = "0"
    UserControl.lblNroBoletas.Caption = "0"
    UserControl.lblNroDocumentos.Caption = "0"
    lblNroAnulados.Caption = "0"
    UserControl.lblTotAnulados.Caption = "0"
    UserControl.lblTotalDevNotaCredito = "0" 'Frank 24082015
    
    'mgaray20101003
    If txtNroSerieBusqueda.Text <> "" And txtNroDocumentoBusqueda.Text <> "" Then
        ms_MensajeError = ms_MensajeError & "NroSerie='" & Trim(txtNroSerieBusqueda.Text) & "' and NroDocumento='" & txtNroDocumentoBusqueda.Text & "' and "
       txtNroHistoriaBusqueda.Text = ""
       mo_cmbIdCajaBusqueda.BoundText = ""
       mo_cmbIdTurnoBusqueda.BoundText = ""
    ElseIf txtNroSerieBusqueda.Text <> "" Then
       ms_MensajeError = ms_MensajeError & "NroSerie='" & Trim(txtNroSerieBusqueda.Text) & "' and "
       txtNroHistoriaBusqueda.Text = ""
       mo_cmbIdCajaBusqueda.BoundText = ""
       mo_cmbIdTurnoBusqueda.BoundText = ""
    ElseIf txtNroDocumentoBusqueda.Text <> "" Then
       ms_MensajeError = ms_MensajeError & "NroDocumento='" & txtNroDocumentoBusqueda.Text & "' and "
       txtNroHistoriaBusqueda.Text = ""
       mo_cmbIdCajaBusqueda.BoundText = ""
       mo_cmbIdTurnoBusqueda.BoundText = ""
    
    ElseIf mo_Teclado.TextoEsSoloNumeros(txtNroHistoriaBusqueda.Text) Then
       ms_MensajeError = ms_MensajeError & "NroHistoriaClinica='" & _
                  Trim(Val(HCigualDNI_AgregaNUEVEaLaHistoria(txtNroHistoriaBusqueda.Text))) & "' and "
       mo_cmbIdCajaBusqueda.BoundText = ""
       mo_cmbIdTurnoBusqueda.BoundText = ""
       txtNroSerieBusqueda.Text = ""
       txtNroDocumentoBusqueda.Text = ""
    ElseIf cmbIdCajaBusqueda.Text <> "" Then
       ms_MensajeError = ms_MensajeError & "idCaja=" & mo_cmbIdCajaBusqueda.BoundText & " and "
       mo_cmbIdTurnoBusqueda.BoundText = ""
       txtNroSerieBusqueda.Text = ""
       txtNroDocumentoBusqueda.Text = ""
       txtNroHistoriaBusqueda.Text = ""
    ElseIf cmbIdTurnoBusqueda.Text <> "" Then
       ms_MensajeError = ms_MensajeError & "idTurno=" & mo_cmbIdTurnoBusqueda.BoundText & " and "
       mo_cmbIdCajaBusqueda.BoundText = ""
       txtNroSerieBusqueda.Text = ""
       txtNroDocumentoBusqueda.Text = ""
       txtNroHistoriaBusqueda.Text = ""
    ElseIf TxtRsocial.Text <> "" Then
       ms_MensajeError = ms_MensajeError & "RazonSocial like '%" & Trim(TxtRsocial.Text) & "%' and "
       mo_cmbIdCajaBusqueda.BoundText = ""
       txtNroSerieBusqueda.Text = ""
       txtNroDocumentoBusqueda.Text = ""
       txtNroHistoriaBusqueda.Text = ""
    End If
    

    
    If cmbIdResponsable.Text <> "" Then
       ms_MensajeError = ms_MensajeError & "idCajero=" & mo_cmbIdResponsable.BoundText & " and "
    End If
    If ms_MensajeError <> "" Then
       ms_MensajeError = Left(ms_MensajeError, Len(ms_MensajeError) - 5)
       oRsBusquedaRecibos.Filter = ms_MensajeError & IIf(chkSoloCredito.Value = 1, " and TieneCredito='C'", " and TieneCredito=null")
       oRsBusquedaDevNotaCredito.Filter = ms_MensajeError 'Frank 24082015
    Else
       oRsBusquedaRecibos.Filter = IIf(chkSoloCredito.Value = 1, "TieneCredito='C'", "TieneCredito=null")
    End If
    
    If oRsBusquedaRecibos.RecordCount > 0 Then
        lnNroBoletas = 0: lnNroFacturas = 0: lnNroDocumentos = 0: lnNroDctosAnulados = 0
        lnTotBoletas = 0: lnTotFacturas = 0: lnTotDctosAnulados = 0: lnTotalRecaudado = 0
        oRsBusquedaRecibos.MoveFirst
        Do While Not oRsBusquedaRecibos.EOF
           If oRsBusquedaRecibos.Fields!idEstadoComprobante = 4 Then
              lnTotalRecaudado = lnTotalRecaudado + oRsBusquedaRecibos.Fields!Total
           ElseIf oRsBusquedaRecibos.Fields!idEstadoComprobante = 6 Then
              lnTotalRecaudado = lnTotalRecaudado - oRsBusquedaRecibos.Fields!Total
           ElseIf oRsBusquedaRecibos.Fields!idEstadoComprobante = 9 Then
              lnNroDctosAnulados = lnNroDctosAnulados + 1
              lnTotDctosAnulados = lnTotDctosAnulados + oRsBusquedaRecibos.Fields!Total
           End If
           'debb-hra-ya en version Polsalud
           If oRsBusquedaRecibos.Fields!IdTipoComprobante = 3 Then
              If oRsBusquedaRecibos.Fields!idEstadoComprobante <> 9 Then
                 lnNroBoletas = lnNroBoletas + 1
                 lnTotBoletas = lnTotBoletas + oRsBusquedaRecibos.Fields!Total
              End If
           ElseIf oRsBusquedaRecibos.Fields!IdTipoComprobante = 2 Then
              If oRsBusquedaRecibos.Fields!idEstadoComprobante <> 9 Then
                 lnNroFacturas = lnNroFacturas + 1
                 lnTotFacturas = lnTotFacturas + oRsBusquedaRecibos.Fields!Total
              End If
           End If
           lnNroDocumentos = lnNroDocumentos + 1
           oRsBusquedaRecibos.MoveNext
        Loop
    End If
    
    'FRANK 24082015
    If oRsBusquedaDevNotaCredito.RecordCount > 0 Then
        lnNroDevNotaCred = 0: lnTotDevNotaCred = 0
        oRsBusquedaDevNotaCredito.MoveFirst
        Do While Not oRsBusquedaDevNotaCredito.EOF
            lnNroDevNotaCred = lnNroDevNotaCred + 1
            lnTotDevNotaCred = lnTotDevNotaCred + oRsBusquedaDevNotaCredito.Fields!Total
            lnNroDocumentos = lnNroDocumentos + 1
            lnTotalRecaudado = lnTotalRecaudado - oRsBusquedaDevNotaCredito.Fields!Total
        oRsBusquedaDevNotaCredito.MoveNext
        Loop
    End If
    
    '''''''''''''''
    txtTotalCajero.Caption = "Recaudado: " & Format(lnTotalRecaudado, "#,###,###.#0")
    'debb-hra-ya en version Polsalud
    UserControl.lblNroDocumentos.Caption = "N° Dctos: " & Format(lnNroDocumentos, "#,###,###")
    UserControl.lblNroFacturas.Caption = "N° Facturas : " & Format(lnNroFacturas, "#,###,###")
    UserControl.lblTotFacturas.Caption = "Tot.Facturas: " & Format(lnTotFacturas, "#,###,###.#0")
    UserControl.lblNroBoletas.Caption = "N° Boletas : " & Format(lnNroBoletas, "#,###,###")
    UserControl.lblTotBoletas.Caption = "Tot.Boletas: " & Format(lnTotBoletas, "#,###,###.#0")
    lblNroAnulados = "N° Anulados : " & Format(lnNroDctosAnulados, "#,###,###")
    UserControl.lblTotAnulados.Caption = "Tot.Anulados: " & Format(lnTotDctosAnulados, "#,###,###.#0")
    'FRANK 24082015
    UserControl.lblNroDevNotaCredito.Caption = "N° NotasCrédito : " & Format(lnNroDevNotaCred, "#,###,###")
    UserControl.lblTotalDevNotaCredito.Caption = "Tot.NotasCrédito: " & Format(lnTotDevNotaCred, "#,###,###.#0")
    ''''''''''''''''
        
    Set grdGestionCaja.DataSource = oRsBusquedaRecibos
    Set grdNotaCredito.DataSource = oRsBusquedaDevNotaCredito 'FRANK 24082015
    ConfigurarGrilla
    'mo_Apariencia.ConfigurarFilasBiColores grdGestionCaja, sighentidades.GrillaConFilasBicolor
    'mo_Apariencia.ConfigurarFilasBiColores grdNotaCredito, sighentidades.GrillaConFilasBicolor 'FRANK 24082015
    ''''''''''''''''
End Sub



'********************************************************************************************
'********************************************************************************************
'********************************************************************************************
'                                   COMPROBANTE DE PAGO
'********************************************************************************************
'********************************************************************************************
'********************************************************************************************

'sunat
Private Sub btnAceptar_Click()
    If btnAceptar.Visible = False Then
       Exit Sub
    End If
 
    lcVuelto = "EFECTIVO: " & txtEfectivo.Text & "        VUELTO: " & txtVuelto.Text
    '
    Dim lIdTipoComprobante As Long
    Dim oDOCajaCaja As DOCajaCaja
    If CCur(txtTotal.Text) < 0 Then
       MsgBox "El total es menor a CERO", vbInformation, "CAJA"
       Exit Sub
    End If
    If mi_Opcion = sghNuevoPagoConHistoria Or mi_Opcion = sghNuevoPagoSinHistoria Then
       lnTotalGrid = ucFacturacionProductos.DevuelveTotalPagar
       ActualizaTotalApagar
    End If
    If MsgBox("Por favor confirmar, ¿Realmente desea grabar los cambios que ha realizado?", vbQuestion + vbYesNo, "Estado de Cuenta") = vbNo Then
        Exit Sub
    End If

    Set oDOCajaCaja = mo_AdminCaja.CajaSeleccionarPorId(mo_DOCajaCaja.IdCaja)
    Select Case mi_Opcion
    'En el caso de las cuenta se va a mostrar Servicios y Bienes e Insumos
    Case sghopcionespago.sghPagarCuentaExistente, sghopcionespago.sghPagarCuentaTotalFS    'debb-17/02/2011
        AgregarComprobantePorCuentaTotal
    Case Else
        If mi_Opcion = sghPagarOrdenExistenteF Then 'solo es usado cuando es "pagar orden existente"
            AgregarComprobanteDeBienes
        Else
            If mi_Opcion = sghopcionespago.sghDevolucion Then
               cmbIdTipoComprobante_Click
               mi_Opcion = sghopcionespago.sghNuevoPagoSinHistoria
            End If
            Select Case ml_TipoProducto
            Case sghServicio
                 If Not IsNull(mo_DOCajaCaja.FormatoImpDefaultCinta) Then
                    lIdTipoComprobante = Val(mo_cmbIdTipoComprobante.BoundText)
                    If lIdTipoComprobante > 0 Then
                        CargaSetup_Caja App.Path & "\archivos", lIdTipoComprobante, oDOCajaCaja.FormatoImpDefaultCinta
                    End If
                 End If
                 AgregarComprobanteDeServicios
            Case sghbien
                 If Not IsNull(mo_DOCajaCaja.FormatoImpDefaultCinta) Then
                    lIdTipoComprobante = Val(mo_cmbIdTipoComprobante.BoundText)
                    If lIdTipoComprobante > 0 Then
                        CargaSetup_Caja App.Path & "\archivos", lIdTipoComprobante, oDOCajaCaja.FormatoImp2Cinta
                    End If
                 End If
                 AgregarComprobanteDeBienes
            End Select
            If optRealizarDevolucion.Value Then
                mi_Opcion = sghopcionespago.sghDevolucion
            End If
        End If
    End Select
    
End Sub



Sub AgregarComprobantePorCuentaTotal()

   Dim lnCuentaActual As Long
   lnCuentaActual = ml_idCuentaAtencion
   If lbCargaEstadoDeCuentaFS = True Then

        If ExisteComprobantePagoPorNroSerieDocumento() = False Then
           '*******************Boletas para Farmacia y Servicio**********************
           '
           '1-Emite Boleta de Farmacia
           txtPagoACuenta.Text = "": txtExonerado.Text = ""
           tabFactProductosPorCuenta.Tab = 1
           mi_Opcion = sghPagarCuentaExistente
           lbCargaEstadoDeCuentaFarmacia = True
           optServicios.Value = False
           txtExonerado.Text = txtCtaFarmExonerado.Text
           lnTotalGrid = Val(txtCtaFarmTfarmacia.Text)
           'CALCULO EN FARMACIA
           CargaDatosDeTotalesDeLaCuenta ml_idCuentaAtencion 'Calcula el pago por adelanto primero por farmacia
           If ValidarDatosObligatorios() Then
                CargaDatosAlObjetosDeDatos
                If AgregarDatosPorCuentaTotal() Then
                      'debb-15/05/2016 (inicio)
                      If lbElNroItemsEsMenorAlMaximoDeBoleta = True Then
                         ImpresionDelRecibo txtNroSerie, txtNroDocumento, sghbien, lnParametrosImprimeBoleta, lbEsUnaFactura
                      End If
                      'debb-15/05/2016 (fin)
                      
                      '2-Emite boleta de Servicios
                      cmbIdTipoComprobante_Click
                      txtPagoACuenta.Text = "": txtExonerado.Text = ""
                      tabFactProductosPorCuenta.Tab = 0
                      mi_Opcion = sghPagarCuentaExistente
                      lbCargaEstadoDeCuentaFarmacia = False
                      optServicios.Value = True
                      txtExonerado.Text = txtCtaServExonerado.Text
                      lnTotalGrid = Val(txtCtaServTservicio.Text)
                      CargaDatosDeTotalesDeLaCuenta ml_idCuentaAtencion
                      If ValidarDatosObligatorios() Then
                           CargaDatosAlObjetosDeDatos
                           If AgregarDatosPorCuentaTotal() Then
                               If lblCuentaConSeguro.Caption = "" Then
                                   'debb-15/05/2016 (inicio)
                                   If lbElNroItemsEsMenorAlMaximoDeBoleta = True Then
                                      ImpresionDelRecibo txtNroSerie, txtNroDocumento, sghServicio, lnParametrosImprimeBoleta, lbEsUnaFactura
                                      cmbIdTipoComprobante_Click
                                      MsgBox "Se emitió Correctamente 2 Boletas", vbInformation, "Caja"
                                   End If
'                                   cmbIdTipoComprobante_Click

'                                   MsgBox "Se emitió Correctamente 2 Boletas", vbInformation, "Caja"
                                   'debb-15/05/2016 (fin)
                               End If
                               '


                               mi_Opcion = sghPagarCuentaTotalFS
                               optPagarCtaTotal.Value = True
                               optPagarCtaTotal_Click True
                               'mgaray201504

                               RaiseEvent GuardoComprobante(True)
                           Else

                               MsgBox "No se pudo agregar los datos (Servicios)" + Chr(13) + ms_MensajeError, vbExclamation, "Gestión de Caja"
                           End If
                      End If
                Else
                      MsgBox "No se pudo agregar los datos (Servicios)" + Chr(13) + ms_MensajeError, vbExclamation, "Gestión de Caja"
                End If
           Else
                  '****emite boleta de Servicios (no hay consumos en farmacia)
                  txtPagoACuenta.Text = "": txtExonerado.Text = ""
                  tabFactProductosPorCuenta.Tab = 0
                  mi_Opcion = sghPagarCuentaExistente
                  lbCargaEstadoDeCuentaFarmacia = False
                  optServicios.Value = True
                  txtExonerado.Text = txtCtaServExonerado.Text
                  lnTotalGrid = Val(txtCtaServTservicio.Text)
                  CargaDatosDeTotalesDeLaCuenta ml_idCuentaAtencion
                  If ValidarDatosObligatorios() Then
                       CargaDatosAlObjetosDeDatos
                       If AgregarDatosPorCuentaTotal() Then
                           If lblCuentaConSeguro.Caption = "" Then
                              'debb-15/05/2016 (inicio)
                              If lbElNroItemsEsMenorAlMaximoDeBoleta = True Then
                                 ImpresionDelRecibo txtNroSerie, txtNroDocumento, sghServicio, lnParametrosImprimeBoleta, lbEsUnaFactura
                                 MsgBox "Se emitió Correctamente", vbInformation, "Caja"
                              End If
                              'debb-15/05/2016 (fin)
                              cmbIdTipoComprobante_Click
                           End If
                           '
                           mi_Opcion = sghPagarCuentaTotalFS
                           optPagarCtaTotal.Value = True
                           optPagarCtaTotal_Click True
                          'mgaray201504
                          RaiseEvent GuardoComprobante(True)
                       Else



                           MsgBox "No se pudo agregar los datos (Servicios)" + Chr(13) + ms_MensajeError, vbExclamation, "Gestión de Caja"
                       End If
                  End If
            End If
















        End If














   Else
        'se usó la cuenta para pagar CONSULTA CE, CONSULTA EMERGENCIA, etc
        If ValidarDatosObligatorios() And ExisteComprobantePagoPorNroSerieDocumento() = False Then
           CargaDatosAlObjetosDeDatos
           If ValidarReglas() Then
               If AgregarDatosPorCuentaTotal() Then
                    If lbCargaEstadoDeCuentaFarmacia = True Then
                       'debb-15/05/2016 (inicio)
                       If lbElNroItemsEsMenorAlMaximoDeBoleta = True Then
                          ImpresionDelRecibo txtNroSerie, txtNroDocumento, sghbien, lnParametrosImprimeBoleta, lbEsUnaFactura
                          MsgBox "Se emitió Correctamente", vbInformation, "Caja"
                       End If
                       'debb-15/05/2016 (fin)

                       optPagarEstadoDeCTAFarmacia_Click 1
                    Else
                       If optServicios.Value Then
                          'debb-15/05/2016 (inicio)
                          If lbElNroItemsEsMenorAlMaximoDeBoleta = True Then
                                ImpresionDelRecibo txtNroSerie, txtNroDocumento, sghServicio, lnParametrosImprimeBoleta, lbEsUnaFactura
                                MsgBox "Se emitió Correctamente", vbInformation, "Caja"
                          End If
                          'debb-15/05/2016 (fin)
                          optPagarEstadoDeCuenta_Click 1
                       Else
                          'debb-15/05/2016 (inicio)
                          If lbElNroItemsEsMenorAlMaximoDeBoleta = True Then
                                ImpresionDelRecibo txtNroSerie, txtNroDocumento, sghbien, lnParametrosImprimeBoleta, lbEsUnaFactura
                                MsgBox "Se emitió Correctamente", vbInformation, "Caja"
                          End If
                          'debb-15/05/2016 (fin)


                          optPagarEstadoDeCTAFarmacia_Click 1
                       End If
                    End If

                    cmbIdTipoComprobante_Click
                    'mgaray201504
                    RaiseEvent GuardoComprobante(True)
                Else
                    MsgBox "No se pudo agregar los datos" + Chr(13) + ms_MensajeError, vbExclamation, "Gestión de Caja"
               End If
           End If
        End If
   End If
   mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar lnCuentaActual, False, 0
End Sub




Sub ImprimirComprobanteServicio()
Dim oRptCaja As New RptCaja
End Sub

Sub ImprimirComprobanteBienesInsumos()
Dim oRptCaja As New RptCaja

    If mo_ReglasComunes.ParametrosSeleccionarValorIntPorTipoYCodigo("INDICADOR", "IMPRIMIR_RECIBO") = 1 Then
        oRptCaja.ImprimirComprobantePagoBienesInsumos mo_DoPaciente, mo_DoAtencion, mo_DOFactOrdenBienInsumo, mo_DOComprobantePago, ucFacturacionProductos.FacturacionProductos
    End If
    
End Sub

'debb-18/05/2016
Function AgregarComprobanteDeServiciosGrabaImprimeXBoleta(oRsItemsDeBoleta As Recordset) As Boolean
        AgregarComprobanteDeServiciosGrabaImprimeXBoleta = False
        If AgregarDatos(oRsItemsDeBoleta) Then
             'cmbOrdenes.Text = mo_DOFactOrdenServicio.IdOrden
             ImpresionDelRecibo txtNroSerie, txtNroDocumento, sghServicio, lnParametrosImprimeBoleta, lbEsUnaFactura
             If mo_DOComprobantePago.idCuentaAtencion > 0 Then
                mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar mo_DOComprobantePago.idCuentaAtencion, False, 0
             End If
             AgregarComprobanteDeServiciosGrabaImprimeXBoleta = True
         Else
             MsgBox "No se pudo emitir DOCUMENTO " + Chr(13) + Chr(13) + ms_MensajeError, vbExclamation, "Gestión de Caja"
         End If

End Function
'debb-18/05/2016

Sub AgregarComprobanteDeServicios()
    Dim lbTodoOKEY As Boolean, lnNumeroDeItemFinal As Long, lnNumeroDeItemInicio As Long
    Dim lnCantidadBoletas As Long, lnExoneracionXBoleta As Double, lnTotalXBoleta As Double
    Dim oRsItemsDeBoleta As New Recordset

    Select Case mi_Opcion
    Case sghNuevoPagoConHistoria, sghNuevoPagoSinHistoria, sghPagarOrdenExistente

            If ValidarDatosObligatorios() And ExisteComprobantePagoPorNroSerieDocumento() = False Then
               CargaDatosAlObjetosDeDatos
               If ValidarReglas() Then
                  lbTodoOKEY = False
                  Set oRsItemsDeBoleta = ucFacturacionProductos.FacturacionProductos.Clone
                  lnCantidadItemsDeLaBoleta = oRsItemsDeBoleta.RecordCount
                  lnCantidadBoletas = 0
                  If wxParametro500 = "S" And lnCantidadItemsDeLaBoleta > Val(wxParametro102) Then
                        '********* imprime VARIAS boletas (inicio)************
                        'lnCantidadBoletas = Round(lnTotalItems / Val(wxParametro102), 0)
                        UserControl.Enabled = False
                        UserControl.MousePointer = 11
                        UserControl.KeyPreview = False
                        
                        lnNumeroDeItemFinal = 1
                        oRsItemsDeBoleta.MoveFirst
                        Do While Not oRsItemsDeBoleta.EOF
                           oRsItemsDeBoleta.Fields!NumeroDeItem = lnNumeroDeItemFinal
                           oRsItemsDeBoleta.Update
                           lnNumeroDeItemFinal = lnNumeroDeItemFinal + 1
                           oRsItemsDeBoleta.MoveNext
                        Loop
                        lnNumeroDeItemInicio = 0
                        lnNumeroDeItemFinal = Val(wxParametro102)
                        Do While True
                           oRsItemsDeBoleta.Filter = "NumeroDeItem>" & lnNumeroDeItemInicio & _
                                                     " and NumeroDeItem<=" & lnNumeroDeItemFinal
                           If oRsItemsDeBoleta.RecordCount = 0 Then
                              Exit Do
                           Else
                              lnTotalXBoleta = 0
                              oRsItemsDeBoleta.MoveFirst
                              Do While Not oRsItemsDeBoleta.EOF
                                   lnTotalXBoleta = lnTotalXBoleta + oRsItemsDeBoleta!TotalPorPagar
                                   oRsItemsDeBoleta.MoveNext
                              Loop
                              lnTotalXBoleta = sighEntidades.DevuelveNumeroRedondeado(lnTotalXBoleta)
                              '
                              lnExoneracionXBoleta = 0
                              If Val(txtExonerado.Text) > 0 Then
                                   If Val(txtExonerado.Text) > lnTotalXBoleta Then
                                      lnExoneracionXBoleta = lnTotalXBoleta
                                      txtExonerado.Text = Val(txtExonerado.Text) - lnTotalXBoleta
                                   Else
                                      lnExoneracionXBoleta = txtExonerado.Text
                                      txtExonerado.Text = 0
                                   End If
                              End If
                              mo_DOComprobantePago.Subtotal = lnTotalXBoleta
                              mo_DOComprobantePago.exoneraciones = lnExoneracionXBoleta
                              mo_DOComprobantePago.Total = lnTotalXBoleta - lnExoneracionXBoleta
                              mo_DoFactOrdenServPagos.ImporteExonerado = lnExoneracionXBoleta
                              '
                              lnNumeroDeItemInicio = lnNumeroDeItemFinal
                              lnNumeroDeItemFinal = lnNumeroDeItemFinal + Val(wxParametro102)
                              lbTodoOKEY = AgregarComprobanteDeServiciosGrabaImprimeXBoleta(oRsItemsDeBoleta)
                              If lbTodoOKEY = False Then
                                 Exit Do
                              End If
                              mo_ReglasComunes.WaitSeconds Val(wxParametro501)
                              lnCantidadBoletas = lnCantidadBoletas + 1
                              cmbIdTipoComprobante_Click
                              mo_DOComprobantePago.nrodocumento = UserControl.txtNroDocumento.Text
                              mo_DOComprobantePago.nroSerie = UserControl.txtNroSerie.Text
                              'mo_DOFactOrdenBienInsumo.IdOrden = 0
                              mo_DoFactOrdenServPagos.IdOrdenPago = 0
                              mo_DOFactOrdenServicio.IdOrden = 0
                           End If
                        Loop
                        
                        UserControl.Enabled = True
                        UserControl.MousePointer = 1
                        UserControl.KeyPreview = True
                        '********* imprime VARIAS boletas (fin)   ************
                  Else
                        '********* imprime UNA boleta - version anterior ************
                          lbTodoOKEY = AgregarComprobanteDeServiciosGrabaImprimeXBoleta(oRsItemsDeBoleta)
                   End If
                   If lbTodoOKEY = True Then
                        MsgBox "Se emitió Correctamente " & IIf(lnCantidadBoletas > 0, Trim(Str(lnCantidadBoletas)) & " Documentos", ""), vbInformation, "Caja"








                        If mi_Opcion = sghNuevoPagoConHistoria Then
                            optNuevoOrdenPagoConHistoria_Click 1
                        ElseIf mi_Opcion = sghNuevoPagoSinHistoria Then
                            optNuevoOrdenPagoSinHistoria_Click 1
                        ElseIf mi_Opcion = sghPagarOrdenExistente Then
                            'debb-16/02/2011
                            optOrdenExistenteFS_Click 1
                            'debb-16/02/2011
                        End If
                        cmbIdTipoComprobante_Click
                       'mgaray201504
                        RaiseEvent GuardoComprobante(True)


                   End If
               End If
           End If
    
    Case sghopcionespago.sghAnulacion
           If ValidarReglas() Then
               If Anulacion() Then
                    MsgBox "La orden se ha anulado correctamente", vbInformation, "Gestión de Caja"
                    optRealizarAnulacion_Click (1)
                    cmbIdTipoComprobante_Click
                Else
                    MsgBox "No se pudo eliminar los datos"
               End If
           End If

    Case sghopcionespago.sghDevolucion
           If ValidarReglas() And ExisteComprobantePagoPorNroSerieDocumento() = False Then
                CargaDatosAlObjetosDeDatos
               If Devolucion() Then
                    ImpresionDelRecibo txtNroSerie, txtNroDocumento, sghServicio, lnParametrosImprimeBoleta, lbEsUnaFactura
                    optRealizarDevolucion_Click (1)
                    cmbIdTipoComprobante_Click
                Else
                    MsgBox "No se pudo eliminar los datos"
               End If
           End If
    Case sghopcionespago.sghReimprimirComprobante
         ImpresionDelRecibo txtNroSerie, txtNroDocumento, sghServicio, lnParametrosImprimeBoleta, lbEsUnaFactura
         cmbIdTipoComprobante_Click
    End Select
    Set oRsItemsDeBoleta = Nothing
    
    UserControl.Enabled = True
    UserControl.MousePointer = 1
    UserControl.KeyPreview = True

End Sub
'debb-18/05/2016
Function AgregarComprobanteDeBienesGrabaImprimeXboleta(oRsItemsDeBoleta As Recordset) As Boolean
    AgregarComprobanteDeBienesGrabaImprimeXboleta = False
    If AgregarDatos(oRsItemsDeBoleta) Then
         ImpresionDelRecibo txtNroSerie, txtNroDocumento, sghbien, lnParametrosImprimeBoleta, lbEsUnaFactura
         
         
         AgregarComprobanteDeBienesGrabaImprimeXboleta = True
     Else
         MsgBox "No se pudo emitir DOCUMENTO " + Chr(13) + ms_MensajeError, vbExclamation, "Gestión de Caja"
    End If
End Function
'debb-18/05/2016

Sub AgregarComprobanteDeBienes()
    Dim lbTodoOKEY As Boolean, lnNumeroDeItemFinal As Long, lnNumeroDeItemInicio As Long
    Dim lnCantidadBoletas As Long, lnExoneracionXBoleta As Double, lnTotalXBoleta As Double, lbEsPrimero As Boolean
    Dim oRsItemsDeBoleta As New Recordset
    Dim oConexion As New Connection
    Dim oFactOrdenesBienes As New FactOrdenesBienes
    Dim oFacturacionBienesPagos As New FacturacionBienesPagos, oDoFacturacionBienesPagos As New DoFacturacionBienesPagos

    Select Case mi_Opcion
    Case sghPagarOrdenExistenteF
        If ValidarDatosObligatorios() And ExisteComprobantePagoPorNroSerieDocumento() = False Then
               CargaDatosAlObjetosDeDatos
               If ValidarReglas() Then
                  lbTodoOKEY = False
                  Set oRsItemsDeBoleta = UcFacturacionContado1.DevuelveProductos
                  lnCantidadItemsDeLaBoleta = oRsItemsDeBoleta.RecordCount
                  lnCantidadBoletas = 0
                  If wxParametro500 = "S" And lnCantidadItemsDeLaBoleta > Val(wxParametro102) Then
                        '********* imprime VARIAS boletas (inicio)************
                        '1- eliminar Items del PRE-PAGO que pasan del Máximo
                        '2- crear Cabecera/detalle del PRE-PAGO de los que pasaron del Máximo
                        '3- emitir Boleta
                        UserControl.Enabled = False
                        UserControl.MousePointer = 11
                        UserControl.KeyPreview = False
                        
                        oConexion.CommandTimeout = 300
                        oConexion.CursorLocation = adUseClient
                        oConexion.Open sighEntidades.CadenaConexion
                        Set oFactOrdenesBienes.Conexion = oConexion
                        Set oFacturacionBienesPagos.Conexion = oConexion
                        
                        lnNumeroDeItemFinal = 1
                        oRsItemsDeBoleta.MoveFirst
                        Do While Not oRsItemsDeBoleta.EOF
                           '
                           If lnNumeroDeItemFinal > Val(wxParametro102) Then
                                oDoFacturacionBienesPagos.IdOrden = mo_DOFactOrdenBienInsumo.IdOrden
                                oDoFacturacionBienesPagos.idProducto = oRsItemsDeBoleta!idProducto
                                oDoFacturacionBienesPagos.IdUsuarioAuditoria = mo_DOFactOrdenBienInsumo.IdUsuarioAuditoria
                                If oFacturacionBienesPagos.EliminarXproducto(oDoFacturacionBienesPagos) = False Then
                                   MsgBox oFacturacionBienesPagos.MensajeError
                                   Exit Do
                                End If
                           End If
                           '
                           oRsItemsDeBoleta.Fields!NumeroDeItem = lnNumeroDeItemFinal
                           oRsItemsDeBoleta.Update
                           lnNumeroDeItemFinal = lnNumeroDeItemFinal + 1
                           oRsItemsDeBoleta.MoveNext
                        Loop
                        lnNumeroDeItemInicio = 0
                        lnNumeroDeItemFinal = Val(wxParametro102)
                        lbEsPrimero = True
                        Do While True
                           oRsItemsDeBoleta.Filter = "NumeroDeItem>" & lnNumeroDeItemInicio & _
                                                     " and NumeroDeItem<=" & lnNumeroDeItemFinal
                           If oRsItemsDeBoleta.RecordCount = 0 Then
                              Exit Do
                           Else
                              If lbEsPrimero = True Then
                              Else
                                 mo_DOFactOrdenBienInsumo.movNumero = ""
                                 mo_DOFactOrdenBienInsumo.MovTipo = ""
                                 mo_DOFactOrdenBienInsumo.IdComprobantePago = 0
                                 If oFactOrdenesBienes.Insertar(mo_DOFactOrdenBienInsumo) = False Then
                                    MsgBox oFactOrdenesBienes.MensajeError
                                    Exit Do
                                 End If
                              End If
                              lnTotalXBoleta = 0
                              oRsItemsDeBoleta.MoveFirst
                              Do While Not oRsItemsDeBoleta.EOF
                                   If lbEsPrimero = True Then
                                   Else
                                      oDoFacturacionBienesPagos.CantidadPagar = oRsItemsDeBoleta!Cantidad
                                      oDoFacturacionBienesPagos.IdOrden = mo_DOFactOrdenBienInsumo.IdOrden
                                      oDoFacturacionBienesPagos.idProducto = oRsItemsDeBoleta!idProducto
                                      oDoFacturacionBienesPagos.IdUsuarioAuditoria = mo_DOFactOrdenBienInsumo.IdUsuarioAuditoria
                                      oDoFacturacionBienesPagos.PrecioVenta = oRsItemsDeBoleta!precio
                                      oDoFacturacionBienesPagos.TotalPagar = oRsItemsDeBoleta!Total
                                      If oFacturacionBienesPagos.Insertar(oDoFacturacionBienesPagos) = False Then
                                         MsgBox oFacturacionBienesPagos.MensajeError
                                         Exit Do
                                      End If
                                   End If
                                   lnTotalXBoleta = lnTotalXBoleta + oRsItemsDeBoleta!Total
                                   oRsItemsDeBoleta.MoveNext
                              Loop
                              lnTotalXBoleta = sighEntidades.DevuelveNumeroRedondeado(lnTotalXBoleta)
                              lbEsPrimero = False
                              '
                              lnExoneracionXBoleta = 0
                              If Val(txtExonerado.Text) > 0 Then
                                   If Val(txtExonerado.Text) > lnTotalXBoleta Then
                                      lnExoneracionXBoleta = lnTotalXBoleta
                                      txtExonerado.Text = Val(txtExonerado.Text) - lnTotalXBoleta
                                   Else
                                      lnExoneracionXBoleta = txtExonerado.Text
                                      txtExonerado.Text = 0
                                   End If
                              End If
                              mo_DOComprobantePago.Subtotal = lnTotalXBoleta
                              mo_DOComprobantePago.exoneraciones = lnExoneracionXBoleta
                              mo_DOComprobantePago.Total = lnTotalXBoleta - lnExoneracionXBoleta
                              mo_DOFactOrdenBienInsumo.ImporteExonerado = lnExoneracionXBoleta
                              '
                              lnNumeroDeItemInicio = lnNumeroDeItemFinal
                              lnNumeroDeItemFinal = lnNumeroDeItemFinal + Val(wxParametro102)
                              lbTodoOKEY = AgregarComprobanteDeBienesGrabaImprimeXboleta(oRsItemsDeBoleta)
                              If lbTodoOKEY = False Then
                                 Exit Do
                              End If
                              mo_ReglasComunes.WaitSeconds Val(wxParametro501)
                              lnCantidadBoletas = lnCantidadBoletas + 1






                              cmbIdTipoComprobante_Click
                              mo_DOComprobantePago.nrodocumento = UserControl.txtNroDocumento.Text
                              mo_DOComprobantePago.nroSerie = UserControl.txtNroSerie.Text
                           End If
                        Loop
                        oConexion.Close
                        UserControl.Enabled = True
                        UserControl.MousePointer = 1
                        UserControl.KeyPreview = True
                        '********* imprime VARIAS boletas (fin)   ************
                  Else
                        '********* imprime UNA boleta - version anterior ************
                        lbTodoOKEY = AgregarComprobanteDeBienesGrabaImprimeXboleta(oRsItemsDeBoleta)
                  End If
                  If lbTodoOKEY = True Then
                      MsgBox "Se emitió Correctamente " & IIf(lnCantidadBoletas > 0, Trim(Str(lnCantidadBoletas)) & " Documentos", ""), vbInformation, "Caja"
                      optOrdenExistenteFS_Click 1
                      cmbIdTipoComprobante_Click
                      RaiseEvent GuardoComprobante(True)
                  End If





               End If
           End If
    Case sghopcionespago.sghAnulacion
           If ValidarReglas() Then
                CargaDatosAlObjetosDeDatos
               If Anulacion() Then
                    MsgBox "La orden se ha anulado correctamente", vbInformation, "Gestión de Caja"
                    optRealizarAnulacion_Click (1)
                    cmbIdTipoComprobante_Click
                Else
                    MsgBox "No se pudo eliminar los datos" + Chr(13) + ms_MensajeError, vbExclamation, "Gestión de Caja"
               End If
           End If
    Case sghopcionespago.sghDevolucion
           If ValidarReglas() And ExisteComprobantePagoPorNroSerieDocumento() = False Then
                CargaDatosAlObjetosDeDatos
               If Devolucion() Then
                    ImpresionDelRecibo txtNroSerie, txtNroDocumento, sghbien, lnParametrosImprimeBoleta, lbEsUnaFactura
                    optRealizarDevolucion_Click (1)
                    cmbIdTipoComprobante_Click
                Else
                    MsgBox "No se pudo eliminar los datos"
               End If
           End If
    Case sghopcionespago.sghReimprimirComprobante
        ImpresionDelRecibo txtNroSerie, txtNroDocumento, sghbien, lnParametrosImprimeBoleta, lbEsUnaFactura
        cmbIdTipoComprobante_Click
    End Select
    Set oRsItemsDeBoleta = Nothing
    Set oConexion = Nothing
    Set oFactOrdenesBienes = Nothing
    Set oFacturacionBienesPagos = Nothing
    Set oDoFacturacionBienesPagos = Nothing
    
    UserControl.Enabled = True
    UserControl.MousePointer = 1
    UserControl.KeyPreview = True

End Sub


Sub LimpiarFormulario()
    Select Case mi_Opcion
    Case sghPagarCuentaExistente, sghPagarCuentaTotalFS    'debb-17/02/2011
        ucFactServiciosPorCuenta.LimpiarGrilla
        ucFactBienesPorCuenta.LimpiarGrilla
        
    Case sghopcionespago.sghPagarOrdenExistente, sghopcionespago.sghPagarOrdenExistenteF, sghopcionespago.sghPagarOrdenExistenteFS 'debb-17/02/2011
        ucFacturacionProductos.FiltraCpt = sghMuestraTodosCpt
        ucFacturacionProductos.LimpiarGrilla
        ucFacturacionProductos.PermiteAgregarItems = False
        UcFacturacionContado1.LimpiarGrilla
        ucFacturacionProductos.LimpiarGrilla
        ucFacturacionProductos.Visible = True
    Case sghopcionespago.sghNuevoPagoConHistoria, sghopcionespago.sghNuevoPagoSinHistoria, sghopcionespago.sghAnulacion, sghopcionespago.sghReimprimirComprobante, sghopcionespago.sghDevolucion
        ucFacturacionProductos.FiltraCpt = sghMuestraTodosCpt
        ucFacturacionProductos.LimpiarGrilla
        If mi_Opcion = sghopcionespago.sghNuevoPagoConHistoria Or mi_Opcion = sghopcionespago.sghNuevoPagoSinHistoria Then
           ucFacturacionProductos.idCuentaAtencion = 0
           ucFacturacionProductos.idTipoFinanciamiento = 1
           ucFacturacionProductos.idPuntoCarga = 99
           ucFacturacionProductos.PermiteAgregarItems = True
           ucFacturacionProductos.AgregaProducto
           If mi_Opcion = sghNuevoPagoSinHistoria And lbEstaCajaUsaDescripcionLarga = True Then
              ucGestionCajaFact1.Visible = True
           End If
        ElseIf mi_Opcion = sghAnulacion Then
           ucGestionCajaFact1.Visible = False
        Else
           ucFacturacionProductos.PermiteAgregarItems = False
        End If
        ucFacturacionProductos.Visible = True
        ucGestionCajaFact1.LimpiarDatos
    End Select
End Sub

Function ValidarDatosObligatorios() As Boolean
    
    ValidarDatosObligatorios = False
    Dim rsProductos As Recordset
    Dim lnParametroParaDevoluciones As Long
    
    If txtNroHistoria.Locked = False Then
        If Trim(txtNroHistoria.Text) = "" Then
            MsgBox "Por favor ingrese el 'Nª de Historia Clínica'", vbInformation, "Caja"
            Exit Function
        End If
    End If
    
    If Val(mo_cmbIdTipoComprobante.BoundText) = 0 Then
        MsgBox "Debe seleccionar el tipo de documento", vbInformation, "Caja"
        Exit Function
    End If
    
    'kike 2017
'    If mi_Opcion = sghNuevoPagoSinHistoria And Val(mo_cmbIdTipoComprobante.BoundText) <> ID_TIPO_COMPROBANTE_FACTURA Then
'        If txtDNI.Text = "" Then
'            MsgBox "Debe ingresar el DNI", vbInformation, "Caja"
'            Exit Function
'        End If
'        If Len(Trim(txtDNI.Text)) <> 8 Then
'            MsgBox "El DNI no tiene el formato correcto", vbInformation, "Caja"
'            Exit Function
'        End If
'    End If
    
    If txtRazonSocial.Text = "" Then
            MsgBox "Por favor ingrese la 'Razón Social'", vbInformation, "Caja"
            txtRazonSocial.SetFocus
            Exit Function
    End If
    
    If txtTotal.Text = "" Then
       txtTotal.Text = "0"
    End If

   
    Select Case mi_Opcion
    Case sghopcionespago.sghPagarCuentaExistente
          If CCur(txtTotal.Text) <= 0 And mi_Opcion <> sghPagarCuentaExistente Then
             MsgBox "El valor TOTAL es menor o igual a cero", vbInformation, "Caja"
             Exit Function
          End If
          If Val(mo_cmbIdTipoComprobante.BoundText) = 2 And Len(Trim(txtRuc.Text)) <> 11 Then
             MsgBox "Debe ingresar el RUC, por favor verifique", vbInformation, "Caja"
             txtRuc.SetFocus
             Exit Function
          End If
          
          If lbCargaEstadoDeCuentaFarmacia = True Then
                'Solo CUENTA FARMACIA en "CAJA DE SERVICIOS"
                If Val(UserControl.txtEfectivo) < Val(UserControl.txtTotal) Then
                    MsgBox "El valor del efectivo no puede ser menor que el monto total, por favor verifique", vbInformation, "Caja"
                    Exit Function
                End If
                Set rsProductos = ucFactBienesPorCuenta.FacturacionProductos
                If Not (rsProductos.EOF And rsProductos.BOF) Then
                    rsProductos.MoveFirst
                    Do While Not rsProductos.EOF
                        If rsProductos!idProducto = 0 Then
                            MsgBox "Uno de los bienes tiene datos imcompletos, por favor verifique", vbInformation, "Caja"
                            Exit Function
                        End If
                        rsProductos.MoveNext
                    Loop
                Else
                   If rsProductos.RecordCount = 0 Then
                           ' MsgBox "No hay items para FARMACIA, por favor verifique", vbInformation, "Caja"
                            Exit Function
                   End If
                End If
                Set rsProductos = UserControl.ucFactServiciosPorCuenta.FacturacionProductos
                If Not (rsProductos.EOF And rsProductos.BOF) Then
                    rsProductos.MoveFirst
                    Do While Not rsProductos.EOF
                        If rsProductos!idProducto = 0 Then
                            MsgBox "Uno de los servicios tiene datos imcompletos, por favor verifique", vbInformation, "Caja"
                            Exit Function
                        End If
                        rsProductos.MoveNext
                    Loop
                End If
          Else
                'Solo CUENTA FARMACIA en "CAJA DE FARMACIA" o CUENTA SERVICIO en "CAJA DE SERVICIO"
                If Val(UserControl.txtEfectivo) < Val(UserControl.txtTotal) Then
                    MsgBox "El valor del efectivo no puede ser menor que el monto total, por favor verifique", vbInformation, "Caja"
                    Exit Function
                End If
                If ml_TipoProducto = sghServicio Then
                   Set rsProductos = ucFactServiciosPorCuenta.FacturacionProductos
                Else
                   Set rsProductos = ucFactBienesPorCuenta.FacturacionProductos
                End If
                If Not (rsProductos.EOF And rsProductos.BOF) Then
                    rsProductos.MoveFirst
                    Do While Not rsProductos.EOF
                        If rsProductos!idProducto = 0 Then
                            MsgBox "Uno de los bienes tiene datos imcompletos, por favor verifique", vbInformation, "Caja"
                            Exit Function
                        End If
                        rsProductos.MoveNext
                    Loop
                Else
                   If rsProductos.RecordCount = 0 Then
                            MsgBox "No hay items para SERVICIOS, por favor verifique", vbInformation, "Caja"
                            Exit Function
                   End If
                End If
                Set rsProductos = UserControl.ucFactServiciosPorCuenta.FacturacionProductos
                If Not (rsProductos.EOF And rsProductos.BOF) Then
                    rsProductos.MoveFirst
                    Do While Not rsProductos.EOF
                        If rsProductos!idProducto = 0 Then
                            MsgBox "Uno de los servicios tiene datos imcompletos, por favor verifique", vbInformation, "Caja"
                            Exit Function
                        End If
                        rsProductos.MoveNext
                    Loop
                End If
         End If
    Case sghNuevoPagoConHistoria, sghNuevoPagoSinHistoria, sghPagarOrdenExistente, sghopcionespago.sghPagarOrdenExistenteF
        If (CCur(txtTotal.Text) + CCur(txtExonerado.Text)) <= 0 Then
             MsgBox "El valor TOTAL es menor o igual a cero", vbInformation, "Caja"
             Exit Function
        End If

                       
        If Val(UserControl.txtEfectivo) < Val(UserControl.txtTotal) Then
            If lbEsDevolucion = False Then
               MsgBox "El valor del efectivo no puede ser menor que el monto total, por favor verifique", vbInformation, "Caja"
               Exit Function
            End If
        End If
        If Val(mo_cmbIdTipoComprobante.BoundText) = 2 And Len(Trim(txtRuc.Text)) <> 11 Then
            MsgBox "Debe ingresar el RUC, por favor verifique", vbInformation, "Caja"
            txtRuc.SetFocus
            Exit Function
        End If
        If mi_Opcion = sghNuevoPagoSinHistoria Or mi_Opcion = sghopcionespago.sghNuevoPagoConHistoria Or _
           mi_Opcion = sghopcionespago.sghPagarOrdenExistenteF Then
           If ucFacturacionProductos.EsUnPagoOtrosAdm = True And txtObservaciones.Text = "" Then
                MsgBox "Debe ingresar el Motivo del INGRESO DEL DINERO en OBSERVACIONES, por favor verifique", vbInformation, "Caja"
                txtObservaciones.SetFocus
                Exit Function
           End If
           If Val(txtExonerado.Text) > 0 And cmbServicioSocial.Text = "" Then
                MsgBox "Debe elegir a la persona de SERVICIO SOCIAL que exoneró", vbInformation, "Caja"
                cmbServicioSocial.Visible = True
                cmbServicioSocial.SetFocus
                Exit Function
           End If
        End If
        If mi_Opcion = sghPagarOrdenExistenteF Then
        Else
            Select Case ml_TipoProducto
            Case sghServicio, sghbien
                lnParametroParaDevoluciones = Val(lcBuscaParametro.SeleccionaFilaParametro(265))
                Set rsProductos = ucFacturacionProductos.FacturacionProductos
                If Not (rsProductos.EOF And rsProductos.BOF) Then
                    rsProductos.MoveFirst
                    Do While Not rsProductos.EOF
                        If rsProductos!idProducto > 0 And rsProductos!idProducto = lnParametroParaDevoluciones And ml_TipoProducto = sghServicio Then
                           lbItemEsDevolucion = True
                        End If
                        '
                        If rsProductos!idProducto = 0 Then
                           rsProductos.Delete
                           rsProductos.Update
                        ElseIf rsProductos.Fields!Cantidad <= 0 Then
                            MsgBox "Tiene problemas en la CANTIDAD, por favor verifique", vbInformation, "Caja"
                            Exit Function
                        ElseIf rsProductos.Fields!PrecioUnitario <= 0 Then
                            MsgBox "Tiene problemas en el PRECIO, por favor verifique", vbInformation, "Caja"
                            Exit Function
                        End If
                        rsProductos.MoveNext
                    Loop
                Else
                   If rsProductos.RecordCount = 0 Then
                            MsgBox "No hay items, por favor verifique", vbInformation, "Caja"
                            Exit Function
                   End If
                End If
            End Select
        End If
    End Select
    
    ValidarDatosObligatorios = True
End Function

Sub CargaDatosAlObjetosDeDatos()
    If txtTotal.Text = "" Then
       txtTotal.Text = "0"
    End If
    Select Case mi_Opcion
    Case sghNuevoPagoConHistoria, sghNuevoPagoSinHistoria
        Select Case ml_TipoProducto
        Case sghServicio
            With mo_DoFactOrdenServPagos
'                 .fechacreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL
                 .fechacreacion = lcBuscaParametro.RetornaFechaServidorSQL & " " & lcBuscaParametro.RetornaHoraServidorSQLserverFormatoGalenhos 'Actualizado FCV 30032015
                 .idestadofacturacion = 4 'Pagado
                 .IdOrden = ml_IdOrdenDespacho
                 .idUsuario = ml_idUsuario
                 .IdUsuarioAuditoria = ml_idUsuario
                 .IdOrdenPago = 0
                 If mi_Opcion = sghopcionespago.sghNuevoPagoSinHistoria Or mi_Opcion = sghopcionespago.sghNuevoPagoConHistoria Then
                    .ImporteExonerado = CCur(txtExonerado.Text)
                    If Val(txtExonerado.Text) > 0 Then
                       .idUsuarioExonera = Val(mo_cmbServicioSocial.BoundText)
                    Else
                       .idUsuarioExonera = 0
                    End If
                 End If
            End With
        Case sghbien
        End Select
    Case sghPagarOrdenExistente, sghopcionespago.sghPagarOrdenExistenteF    'debb-04/09/2018
        If mi_Opcion = sghPagarOrdenExistenteF Then
            With mo_DOFactOrdenBienInsumo
                 .idestadofacturacion = 4   'pagado
                 .IdUsuarioAuditoria = ml_idUsuario
                 If Val(txtExonerado.Text) > 0 Then
                   .idUsuarioExonera = Val(mo_cmbServicioSocial.BoundText)
                   .ImporteExonerado = CCur(txtExonerado.Text)
                 Else
                   .idUsuarioExonera = 0
                   .ImporteExonerado = 0
                 End If
            End With
        Else
            Select Case ml_TipoProducto
            Case sghServicio
                With mo_DoFactOrdenServPagos
                    .idestadofacturacion = 4   'pagado
                    .IdUsuarioAuditoria = ml_idUsuario
                End With
            Case sghbien
                With mo_DOFactOrdenBienInsumo
                     .idestadofacturacion = 4   'pagado
                     .IdUsuarioAuditoria = ml_idUsuario
                End With
            End Select
        End If
    End Select

    Select Case mi_Opcion
    Case sghNuevoPagoConHistoria, sghNuevoPagoSinHistoria, sghopcionespago.sghPagarCuentaExistente, _
           sghPagarOrdenExistente, sghopcionespago.sghPagarOrdenExistenteF

        Set mo_DOComprobantePago = New DOCajaComprobantesPago
        With mo_DOComprobantePago
            .IdTipoComprobante = Val(mo_cmbIdTipoComprobante.BoundText)
            .nroSerie = Trim(txtNroSerie.Text)
            .nrodocumento = Trim(txtNroDocumento.Text)
            .idCuentaAtencion = ml_idCuentaAtencion
            .razonSocial = txtRazonSocial.Text
            .Observaciones = txtObservaciones.Text
            .IdGestionCaja = mo_doCajaGestion.IdGestionCaja
            .IdUsuarioAuditoria = ml_idUsuario
            .ruc = txtRuc.Text
            .DNI = IIf(txtDni.Text = "", lcDNIbuscado, txtDni.Text)
            .Subtotal = CCur(txtTotal.Text) + CCur(txtExonerado.Text) + CCur(txtPagoACuenta.Text)
            .IGV = 0
            .Total = CCur(txtTotal.Text)
            .FechaCobranza = lcBuscaParametro.RetornaFechaServidorSQL & " " & lcBuscaParametro.RetornaHoraServidorSQLserverFormatoGalenhos 'Actualizado FCV 30032015
            .fechaEmision = .FechaCobranza
            .IdComprobantePago = 0
            If lbEsDevolucion Then
               .IdTipoPago = 2 'devolucion
            Else
               .IdTipoPago = 1 'Orden de pago
            End If

            .idPaciente = ml_IdPaciente
            If lnIdPacienteDelDNIelegido > 0 And ml_IdPaciente = 0 Then
               .idPaciente = lnIdPacienteDelDNIelegido
            End If
            
            .IdFormaPago = ml_IdFormaPago
            .idFarmacia = ml_IdFarmacia
            .IdCaja = Val(mo_cmbIdCaja.BoundText)
            .IdTurno = Val(mo_cmbIdTurno.BoundText)
            .IdCajero = ml_idUsuario
            .IdUsuarioAuditoria = ml_idUsuario
            .Dctos = 0
            .exoneraciones = CCur(txtExonerado.Text)
            .Adelantos = CCur(txtPagoACuenta.Text)
            .idTipoFinanciamiento = Val(mo_cmbIdTipoFinanciamiento.BoundText)
        End With
        If lbItemEsDevolucion = True Then
           mo_DOComprobantePago.IdTipoPago = 2
        End If
    Case sghopcionespago.sghDevolucion
        With mo_DOComprobantePagoDevolucion
            .IdTipoComprobante = Val(mo_cmbIdTipoComprobante.BoundText)
            .nroSerie = Trim(txtNroSerie.Text)
            .nrodocumento = Trim(txtNroDocumento.Text)
            .idCuentaAtencion = mo_DoAtencion.idCuentaAtencion
            .razonSocial = txtRazonSocial.Text
            .Observaciones = ""
            .IdGestionCaja = mo_doCajaGestion.IdGestionCaja
            .IdUsuarioAuditoria = ml_idUsuario
            .ruc = txtRuc.Text
            .Subtotal = CCur(txtTotal)
            .IGV = 0
            .Total = CCur(txtTotal)
'            .FechaCobranza = lcBuscaParametro.RetornaFechaHoraServidorSQL    ' Now
            .FechaCobranza = lcBuscaParametro.RetornaFechaServidorSQL & " " & lcBuscaParametro.RetornaHoraServidorSQLserverFormatoGalenhos 'Actualizado FCV 30032015
            .IdComprobantePago = 0
            .IdTipoPago = 2 'devolucion
            .IdTipoOrden = mo_DOComprobantePago.IdTipoOrden
        End With
    
    End Select
    CargaDatosAlObjetosDeDatosFactura
End Sub

Function ValidarReglas() As Boolean
   ValidarReglas = False
      
   If mi_Opcion = sghPagarOrdenExistenteF Then
      txtNreceta.Text = ""
   End If

   ValidarReglas = True
End Function



'debb-18/05/2016
Function AgregarDatos(oRsItems As Recordset) As Boolean
    Dim oDllFactUCGestionCaja As New SighFacturacion.dllFactUCGestionCaja
    If mi_Opcion = sghPagarOrdenExistenteF Then 'solo es usado cuando es "pagar orden existente"
        '
        AgregarDatos = oDllFactUCGestionCaja.CajaComprobantePagoBienesRegistraBoleta(mo_DOComprobantePago, _
                            mo_doCajaGestion, mo_DOFactOrdenBienInsumo, oRsItems, _
                            ml_idUsuario, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtRazonSocial.Text)
        '
        ms_MensajeError = oDllFactUCGestionCaja.MensajeError    'mo_ReglasFarmacia.MensajeError
    Else
        Select Case ml_TipoProducto
        Case sghServicio
            If lnIdFactPaquete > 0 Then
               AgregarDatos = mo_AdminCaja.CajaComprobantePagoServicioPaqueteAgregar(mo_DOComprobantePago, _
                                           mo_doCajaGestion, mo_DoFactOrdenServPagos, _
                                           oRsItems, ml_idUsuario, mo_DoAtencion, _
                                           Val(mo_cmbIdPuntoCarga.BoundText), mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
               ms_MensajeError = mo_AdminCaja.MensajeError
            Else
               AgregarDatos = oDllFactUCGestionCaja.CajaComprobantePagoServicioAgregar(mo_DOComprobantePago, _
                                            mo_doCajaGestion, mo_DoFactOrdenServPagos, _
                                            oRsItems, ml_idUsuario, _
                                            mo_DoAtencion, Val(mo_cmbIdPuntoCarga.BoundText), _
                                            mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lnIdReceta, lbEsUnaRecetaOtrosCpt)
               ms_MensajeError = oDllFactUCGestionCaja.MensajeError
            End If
        Case sghbien
            AgregarDatos = mo_AdminCaja.CajaComprobantePagoBienInsumoAgregar(mo_DOComprobantePago, mo_doCajaGestion, _
                                            mo_DOFactOrdenBienInsumo, ucFacturacionProductos.FacturacionProductos, _
                                            ucFacturacionProductos.ProductosEliminados, ml_idUsuario, _
                                            mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
            ms_MensajeError = mo_AdminCaja.MensajeError
        End Select
        
    End If
    Set oDllFactUCGestionCaja = Nothing
End Function

'debb-18/05/2016
Function AgregarDatosPorCuentaTotalSeCierraCta(rs1 As Recordset, oConexion As Connection, lbEsBoletaFarmacia As Boolean) As Long
        Dim rs As Recordset
        AgregarDatosPorCuentaTotalSeCierraCta = ml_idCuentaAtencion
        If rs1.Fields!idTipoServicio = 3 And _
            mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(rs1.Fields!IdFormaPago, oConexion) = True Then
            If lbEsBoletaFarmacia = True Then
                'Servicios Pendientes por Pagar para ver si se CIERRA LA CUENTA
                Set rs = mo_ReglasFacturacion.FactOrdenServicioFiltraPorIdCuenta(ml_idCuentaAtencion)
            Else
                'Medicamentos Pendientes por Pagar para ver si se CIERRA LA CUENTA
                Set rs = mo_ReglasFacturacion.FactOrdenesBienesSeleccionarPorIdCuenta(ml_idCuentaAtencion)
            End If
            If rs.RecordCount > 0 Then
               rs.MoveFirst
               Do While Not rs.EOF
                  If rs.Fields!idestadofacturacion = 1 Then
                     AgregarDatosPorCuentaTotalSeCierraCta = 0

                     Exit Do
                  End If
                  rs.MoveNext
               Loop
            End If
            rs.Close
        Else
             AgregarDatosPorCuentaTotalSeCierraCta = 0

        End If
        Set rs = Nothing
End Function

'debb-18/05/2016
'Cuenta - Paciente con alta medica - Sis,Soat,etc
'Puede tener varios PUNTOS DE CARGA (Laboratorio,imagenes,otros) y se tiene q emitir BOLETA a cada uno
Sub CuentaConAlgunSeguroConItemsQueTieneQuePagar(lnIdCuentaAtencionAcerrar As Long, _
                                                 oDllFactUCGestionCaja As SighFacturacion.dllFactUCGestionCaja, _
                                                 ByRef AgregarDatosPorCuentaTotal As Boolean)
                                                
               Dim oRsDetalleBoleta As New Recordset
               Dim oRsSoloCuentasSeguros As New Recordset
               Dim lnIdPuntoCarga As Long
               Dim oCajaNroDocumento As New DOCajaNroDocumento
               Set oRsDetalleBoleta = ucFactServiciosPorCuenta.FacturacionProductos
               With oRsSoloCuentasSeguros
                      .Fields.Append "idPuntoCarga", adInteger
                      .Fields.Append "Total", adDouble
                      .CursorType = adOpenDynamic
                      .LockType = adLockOptimistic
                      .Open
               End With
               oRsDetalleBoleta.MoveFirst
               Do While Not oRsDetalleBoleta.EOF
                  lnIdPuntoCarga = IIf(IsNull(oRsDetalleBoleta.Fields!idPuntoCarga), 0, oRsDetalleBoleta.Fields!idPuntoCarga)
                  If oRsSoloCuentasSeguros.RecordCount > 0 Then
                     oRsSoloCuentasSeguros.MoveFirst
                     oRsSoloCuentasSeguros.Find "idPuntoCarga=" & lnIdPuntoCarga
                     If oRsSoloCuentasSeguros.EOF Then
                        oRsSoloCuentasSeguros.AddNew
                        oRsSoloCuentasSeguros.Fields!idPuntoCarga = lnIdPuntoCarga
                        oRsSoloCuentasSeguros.Fields!Total = oRsDetalleBoleta.Fields!TotalPorPagar
                        oRsSoloCuentasSeguros.Update
                     Else
                        oRsSoloCuentasSeguros.Fields!Total = oRsSoloCuentasSeguros.Fields!Total + oRsDetalleBoleta.Fields!TotalPorPagar
                        oRsSoloCuentasSeguros.Update
                     End If
                  Else
                     oRsSoloCuentasSeguros.AddNew
                     oRsSoloCuentasSeguros.Fields!idPuntoCarga = lnIdPuntoCarga
                     oRsSoloCuentasSeguros.Fields!Total = oRsDetalleBoleta.Fields!TotalPorPagar
                     oRsSoloCuentasSeguros.Update
                  End If
                  oRsDetalleBoleta.MoveNext
               Loop
               If oRsSoloCuentasSeguros.RecordCount = 1 Then
                  lblCuentaConSeguro.Caption = ""
                  AgregarDatosPorCuentaTotal = oDllFactUCGestionCaja.CajaComprobantePagoCuentaAtencionAgregarS(mo_DOComprobantePago, _
                                                mo_doCajaGestion, ucFactServiciosPorCuenta.FacturacionProductos, _
                                                lnIdCuentaAtencionAcerrar, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
               Else
                  oRsSoloCuentasSeguros.MoveFirst
                  Do While Not oRsSoloCuentasSeguros.EOF
                     oRsDetalleBoleta.Filter = "idPuntoCarga=" & oRsSoloCuentasSeguros.Fields!idPuntoCarga
                     mo_DOComprobantePago.nrodocumento = UserControl.txtNroDocumento.Text
                     mo_DOComprobantePago.nroSerie = UserControl.txtNroSerie.Text
                     mo_DOComprobantePago.Subtotal = oRsSoloCuentasSeguros.Fields!Total
                     mo_DOComprobantePago.Total = oRsSoloCuentasSeguros.Fields!Total
                     mo_doCajaGestion.TotalCobrado = oRsSoloCuentasSeguros.Fields!Total
                     '
                     AgregarDatosPorCuentaTotal = oDllFactUCGestionCaja.CajaComprobantePagoCuentaAtencionAgregarS(mo_DOComprobantePago, _
                                                  mo_doCajaGestion, oRsDetalleBoleta, lnIdCuentaAtencionAcerrar, _
                                                  mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
                     '
                     ImpresionDelRecibo txtNroSerie, txtNroDocumento, sghServicio, lnParametrosImprimeBoleta, lbEsUnaFactura
                     Set oCajaNroDocumento = mo_AdminCaja.NroDocumentoSeleccionarPorIdCajaYTipoComprobante(mo_doCajaGestion.IdCaja, Val(mo_cmbIdTipoComprobante.BoundText))
                     txtNroSerie.Text = Trim(oCajaNroDocumento.nroSerie)
                     txtNroDocumento.Text = Trim(oCajaNroDocumento.nrodocumento)
                     oRsSoloCuentasSeguros.MoveNext
                  Loop
                  MsgBox "Se emitió correctamente " & Trim(Str(oRsSoloCuentasSeguros.RecordCount)) & " Boletas"
               End If
               
               oRsDetalleBoleta.Close
               oRsSoloCuentasSeguros.Close
               Set oRsDetalleBoleta = Nothing
               Set oRsSoloCuentasSeguros = Nothing
               Set oCajaNroDocumento = Nothing
               AgregarDatosPorCuentaTotal = True
End Sub
'debb-18/05/2016
Function AgregarDatosPorCuentaTotal() As Boolean
Dim lnIdCuentaAtencionAcerrar As Long, lcEstadosFacturacion As String, lcTiposFinanciamiento As String, lnIdTipoServicio As Long
Dim oDllFactUCGestionCaja As New SighFacturacion.dllFactUCGestionCaja
Dim rs As Recordset
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion

    lnIdCuentaAtencionAcerrar = ml_idCuentaAtencion
    Set rs = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(ml_idCuentaAtencion, oConexion)

    If lbCargaEstadoDeCuentaFarmacia = True Then
        'Cuenta - Paciente Pagante - solo Farmacia
        'Documento de MEDICAMENTOS emitido en Caja Servicio
        AgregarDatosPorCuentaTotal = AgregarDatosPorCuentaTotalGrabaBoletaBS(True, rs, oConexion, _
                                                                             ucFactBienesPorCuenta.FacturacionProductos, 3, _
                                                                             lnIdCuentaAtencionAcerrar, oDllFactUCGestionCaja)

    Else
        If optServicios.Value Then
            If lblCuentaConSeguro.Caption = "" Then
               'Cuenta - Paciente Pagante - solo Servicios
               'Documento de SERVICIOS emitido en Caja Servicio
               AgregarDatosPorCuentaTotal = AgregarDatosPorCuentaTotalGrabaBoletaBS(False, rs, oConexion, _
                                                                                  ucFactServiciosPorCuenta.FacturacionProductos, 1, _
                                                                                  lnIdCuentaAtencionAcerrar, oDllFactUCGestionCaja)
            Else
                'Cuenta - Paciente con alta medica - Paciente con SEGUROS que algunos items no lo cubre
               lnIdCuentaAtencionAcerrar = AgregarDatosPorCuentaTotalSeCierraCta(rs, oConexion, False)
               mo_DOComprobantePago.IdTipoOrden = 1  'Documento de SERVICIOS emitido en Caja Servicio
               CuentaConAlgunSeguroConItemsQueTieneQuePagar lnIdCuentaAtencionAcerrar, oDllFactUCGestionCaja, AgregarDatosPorCuentaTotal
            End If
        Else
            'items de FARMACIA
            'Documento de MEDICAMENTOS emitido en Caja Farmacia
            AgregarDatosPorCuentaTotal = AgregarDatosPorCuentaTotalGrabaBoletaBS(True, rs, oConexion, _
                                                                                ucFactBienesPorCuenta.FacturacionProductos, 2, _
                                                                                lnIdCuentaAtencionAcerrar, oDllFactUCGestionCaja)
        End If
    End If
    ms_MensajeError = mo_AdminCaja.MensajeError
    Set rs = Nothing
    Set oDllFactUCGestionCaja = Nothing
    oConexion.Close
    Set oConexion = Nothing
End Function
'debb-18/05/2016
Function AgregarDatosPorCuentaTotalGrabaBoletaBS(lbEsBoletaFarmacia As Boolean, rs As Recordset, oConexion As Connection, _
                                         oRsItemsCuenta As Recordset, lnIdTipoOrden As Long, _
                                         lnIdCuentaAtencionAcerrar As Long, _
                                         oDllFactUCGestionCaja As SighFacturacion.dllFactUCGestionCaja) As Boolean
    On Error GoTo ErrAgrDtosCtaTotGrab
    Dim oFactOrdenesBienes As New FactOrdenesBienes, oDoFactOrdenesBienes As New DoFactOrdenesBienes
    Dim oFacturacionBienesPagos As New FacturacionBienesPagos, oDoFacturacionBienesPagos As New DoFacturacionBienesPagos
    Dim oFactOrdenServicioPagos As New FactOrdenServicioPagos, oDoFactOrdenServPagos As New DoFactOrdenServPagos
    Dim oFacturacionServicioPagos As New FacturacionServicioPagos, oDoFacturacionServicioPagos As New DoFacturacionServicioPagos
    Dim oRsExoneracionesTmp As New Recordset, oRsTmp2 As New Recordset, oRsOrdenPagoTmp As New Recordset
    Dim lnNroItemsXordenPago As Integer, lnNumeroDeItemFinal As Integer, lbNuevoOrdenPago As Boolean, lnIdOrdenPago000 As Long
    Dim lnImporteTotal111 As Double, lnIdOrdenPago111 As Long, lnPagoACuenta As Double, lnAdelantos111 As Double
    Dim lnImporteExonerado111  As Double, lnIdUsuarioExonera111 As Long, lnIdOrden888 As Long, lcMovNumero888 As String
    Dim lcError As String
    UserControl.Enabled = False
    UserControl.MousePointer = 11
    UserControl.KeyPreview = False
     
    
    lcError = ""
    lnIdCuentaAtencionAcerrar = AgregarDatosPorCuentaTotalSeCierraCta(rs, oConexion, lbEsBoletaFarmacia)
    mo_DOComprobantePago.IdTipoOrden = lnIdTipoOrden    '1 (Documento de SERVICIOS emitido en Caja Servicio)
                                                        '2 (Documento de MEDICAMENTOS emitido en Caja Farmacia)
                                                        '3 (Documento de MEDICAMENTOS emitido en Caja Servicio)
    lnCantidadItemsDeLaBoleta = oRsItemsCuenta.RecordCount
    If wxParametro500 = "S" And lnCantidadItemsDeLaBoleta > Val(wxParametro102) Then
    
            lbElNroItemsEsMenorAlMaximoDeBoleta = False
            '******************************************************* imprime VARIAS boletas (inicio)************
            '1- Chequear que los Items de cada IdordenPago no pasen del Máximo,
            '            si pasan dividirlo en 2 o más IdOrdenPago
            '2- Generar Boleta para cada IdOrdenPago
            '3- emitir Boleta para cada IdOrdenPago
            With oRsOrdenPagoTmp
                  .Fields.Append "IdOrdenPago", adInteger
                  .Fields.Append "ImporteExonerado", adDouble
                  .Fields.Append "Adelantos", adDouble
                  .Fields.Append "SubTotal", adDouble
                  .Fields.Append "idUsuarioExonera", adInteger
                  .Fields.Append "IdOrdenPagoAnterior", adInteger
                  .LockType = adLockOptimistic
                  .Open
            End With
            
            If lbEsBoletaFarmacia = True Then
                'Farmacia
                Set oRsExoneracionesTmp = mo_ReglasFacturacion.FacturacionBienesFinanciamientosXcuenta(ml_idCuentaAtencion, oConexion)
                Set oFacturacionBienesPagos.Conexion = oConexion
                Set oFactOrdenesBienes.Conexion = oConexion
                oRsItemsCuenta.Sort = "idOrden"
                oRsItemsCuenta.MoveFirst
                Do While Not oRsItemsCuenta.EOF
                   lnIdOrdenPago000 = oRsItemsCuenta!IdOrden
                   lnNumeroDeItemFinal = 1
                   lnNroItemsXordenPago = 1
                   oDoFactOrdenesBienes.IdUsuarioAuditoria = mo_DOComprobantePago.IdUsuarioAuditoria
                   oDoFactOrdenesBienes.IdOrden = lnIdOrdenPago000
                   If oFactOrdenesBienes.SeleccionarPorId(oDoFactOrdenesBienes) = False Then
                       GoTo ErrAgrDtosCtaTotGrab: lcError = oFactOrdenesBienes.MensajeError
                   End If
                   lcMovNumero888 = oDoFactOrdenesBienes.movNumero
                   Do While Not oRsItemsCuenta.EOF And lnIdOrdenPago000 = oRsItemsCuenta!IdOrden
                        If lnNumeroDeItemFinal > Val(wxParametro102) Then
                                oDoFacturacionBienesPagos.IdOrden = lnIdOrdenPago000
                                oDoFacturacionBienesPagos.idProducto = oRsItemsCuenta!idProducto
                                oDoFacturacionBienesPagos.IdUsuarioAuditoria = mo_DOComprobantePago.IdUsuarioAuditoria
                                If oFacturacionBienesPagos.EliminarXproducto(oDoFacturacionBienesPagos) = False Then
                                   GoTo ErrAgrDtosCtaTotGrab: lcError = oFacturacionBienesPagos.MensajeError
                                End If
                                If lnNroItemsXordenPago > Val(wxParametro102) Then
                                     If oFactOrdenesBienes.Insertar(oDoFactOrdenesBienes) = False Then
                                        GoTo ErrAgrDtosCtaTotGrab: lcError = oFactOrdenesBienes.MensajeError
                                     End If
                                     lnNroItemsXordenPago = 1
                                End If
                                oDoFacturacionBienesPagos.CantidadPagar = oRsItemsCuenta!Cantidad
                                oDoFacturacionBienesPagos.IdOrden = oDoFactOrdenesBienes.IdOrden
                                oDoFacturacionBienesPagos.idProducto = oRsItemsCuenta!idProducto
                                oDoFacturacionBienesPagos.IdUsuarioAuditoria = mo_DOComprobantePago.IdUsuarioAuditoria
                                oDoFacturacionBienesPagos.PrecioVenta = oRsItemsCuenta!PrecioUnitario
                                oDoFacturacionBienesPagos.TotalPagar = oRsItemsCuenta!TotalPorPagar
                                If oFacturacionBienesPagos.Insertar(oDoFacturacionBienesPagos) = False Then
                                   GoTo ErrAgrDtosCtaTotGrab: lcError = oFacturacionBienesPagos.MensajeError
                                End If
                                lnIdOrdenPago111 = oDoFactOrdenesBienes.IdOrden
                        Else
                                lnIdOrdenPago111 = oRsItemsCuenta!IdOrden
                        End If
                        lnImporteTotal111 = oRsItemsCuenta!TotalPorPagar
                        'actualiza EXONERACIONES, subTotal
                        lnImporteExonerado111 = 0
                        lnIdUsuarioExonera111 = 0
                        oRsExoneracionesTmp.Filter = "movnumero='" & lcMovNumero888 & "'" & _
                                                     " and idProducto=" & oRsItemsCuenta!idProducto & _
                                                     " and IdTipoFinanciamiento = 9 and IdEstadoFacturacion <> 9"
                        If oRsExoneracionesTmp.RecordCount > 0 Then
                           lnImporteExonerado111 = oRsExoneracionesTmp!TotalFinanciado
                           lnIdUsuarioExonera111 = oRsExoneracionesTmp!IdUsuarioAutoriza
                        End If
                        lbNuevoOrdenPago = True
                        If oRsOrdenPagoTmp.RecordCount > 0 Then
                           oRsOrdenPagoTmp.MoveFirst
                           oRsOrdenPagoTmp.Find "idOrdenPago=" & lnIdOrdenPago111
                           If Not oRsOrdenPagoTmp.EOF Then
                              lbNuevoOrdenPago = False
                           End If
                        End If
                        If lbNuevoOrdenPago = True Then
                           oRsOrdenPagoTmp.AddNew
                           oRsOrdenPagoTmp!IdOrdenPago = lnIdOrdenPago111
                           oRsOrdenPagoTmp!ImporteExonerado = lnImporteExonerado111
                           oRsOrdenPagoTmp!Subtotal = lnImporteTotal111
                           oRsOrdenPagoTmp!IdOrdenPagoAnterior = lnIdOrdenPago000
                        Else
                           oRsOrdenPagoTmp!ImporteExonerado = oRsOrdenPagoTmp!ImporteExonerado + lnImporteExonerado111
                           oRsOrdenPagoTmp!Subtotal = oRsOrdenPagoTmp!Subtotal + lnImporteTotal111
                        End If
                        If lnIdUsuarioExonera111 > 0 Then
                           oRsOrdenPagoTmp!idUsuarioExonera = lnIdUsuarioExonera111
                        End If
                        oRsOrdenPagoTmp.Update
                        '
                        oRsItemsCuenta.Fields!idordenPagoNEW = lnIdOrdenPago111
                        oRsItemsCuenta.Update
                        oRsItemsCuenta.MoveNext
                        lnNumeroDeItemFinal = lnNumeroDeItemFinal + 1
                        lnNroItemsXordenPago = lnNroItemsXordenPago + 1
                        If oRsItemsCuenta.EOF Then
                           Exit Do
                        End If
                   Loop
                   'actualiza exoneraciones, solo para aquellos que pasan del Máximo de Items
                   If (lnNumeroDeItemFinal - 1) > Val(wxParametro102) Then
                      oRsExoneracionesTmp.Filter = "movNumero='" & lcMovNumero888 & "'" & _
                                                   " and IdTipoFinanciamiento = 9 and IdEstadoFacturacion <> 9"
                      If oRsExoneracionesTmp.RecordCount > 0 Then
                         oRsOrdenPagoTmp.Filter = "IdOrdenPagoAnterior=" & lnIdOrdenPago000
                         If oRsOrdenPagoTmp.RecordCount > 0 Then
                            oRsOrdenPagoTmp.MoveFirst
                            Do While Not oRsOrdenPagoTmp.EOF
                               If mo_ReglasFacturacion.FactOrdenesBienesActualizaExoneracion(oRsOrdenPagoTmp!IdOrdenPago, _
                                                                         oRsOrdenPagoTmp!ImporteExonerado, _
                                                                         oRsOrdenPagoTmp!idUsuarioExonera, oConexion) = False Then
                                  GoTo ErrAgrDtosCtaTotGrab: lcError = mo_ReglasFacturacion.MensajeError
                               End If
                               oRsOrdenPagoTmp.MoveNext
                            Loop
                         End If
                      End If
                   End If
                   '
                Loop
                oRsItemsCuenta.Sort = ""
                oRsItemsCuenta.MoveFirst
                Do While Not oRsItemsCuenta.EOF
                   oRsItemsCuenta!IdOrden = oRsItemsCuenta!idordenPagoNEW
                   oRsItemsCuenta.Update
                   oRsItemsCuenta.MoveNext
                Loop
                'graba e imprime cada Boleta por cada RECETA, que ya no se parasan del MAXIMO DE ITEMS
                lnPagoACuenta = CCur(txtPagoACuenta.Text)
                oRsOrdenPagoTmp.Filter = ""
                oRsOrdenPagoTmp.MoveFirst
                Do While Not oRsOrdenPagoTmp.EOF
                    If (oRsOrdenPagoTmp!Subtotal - (oRsOrdenPagoTmp!ImporteExonerado + lnPagoACuenta)) <= 0 Then
                       lnAdelantos111 = oRsOrdenPagoTmp!Subtotal - oRsOrdenPagoTmp!ImporteExonerado
                       lnPagoACuenta = lnPagoACuenta - lnAdelantos111
                    Else
                       lnAdelantos111 = lnPagoACuenta
                       lnPagoACuenta = 0
                    End If
                    mo_DOComprobantePago.exoneraciones = oRsOrdenPagoTmp!ImporteExonerado
                    mo_DOComprobantePago.Adelantos = lnAdelantos111
                    mo_DOComprobantePago.Subtotal = oRsOrdenPagoTmp!Subtotal
                    mo_DOComprobantePago.Total = oRsOrdenPagoTmp!Subtotal - (oRsOrdenPagoTmp!ImporteExonerado + lnAdelantos111)
                    oRsItemsCuenta.Filter = "idOrdenPagoNew=" & oRsOrdenPagoTmp!IdOrdenPago
                    If oRsItemsCuenta.RecordCount > 0 Then
                        AgregarDatosPorCuentaTotalGrabaBoletaBS = oDllFactUCGestionCaja.CajaComprobantePagoCuentaAtencionAgregarB(mo_DOComprobantePago, _
                                                 oRsItemsCuenta, lnIdCuentaAtencionAcerrar, _
                                                 mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
                        If AgregarDatosPorCuentaTotalGrabaBoletaBS = True Then
                            ImpresionDelRecibo txtNroSerie, txtNroDocumento, sghbien, lnParametrosImprimeBoleta, lbEsUnaFactura
                            mo_ReglasComunes.WaitSeconds Val(wxParametro501)
                            cmbIdTipoComprobante_Click
                            mo_DOComprobantePago.nrodocumento = UserControl.txtNroDocumento.Text
                            mo_DOComprobantePago.nroSerie = UserControl.txtNroSerie.Text
                        Else
                            lcError = lcError & " falta Boleta para ORDEN DE PAGO: " & Trim(Str(oRsOrdenPagoTmp!IdOrdenPago)) & Chr(13)
                        End If
                    End If
                    oRsOrdenPagoTmp.MoveNext
                Loop
            Else
                'servicios
                Set oRsExoneracionesTmp = mo_ReglasFacturacion.FacturacionServicioFinanciamientosXcuentaConexion(ml_idCuentaAtencion, oConexion)
                Set oFacturacionServicioPagos.Conexion = oConexion
                Set oFactOrdenServicioPagos.Conexion = oConexion
                oRsItemsCuenta.Sort = "idOrden"
                oRsItemsCuenta.MoveFirst
                Do While Not oRsItemsCuenta.EOF
                   lnIdOrdenPago000 = oRsItemsCuenta!IdOrden
                   lnNumeroDeItemFinal = 1
                   lnNroItemsXordenPago = 1
                   oDoFactOrdenServPagos.IdUsuarioAuditoria = mo_DOComprobantePago.IdUsuarioAuditoria
                   oDoFactOrdenServPagos.IdOrdenPago = lnIdOrdenPago000
                   If oFactOrdenServicioPagos.SeleccionarPorId(oDoFactOrdenServPagos) = False Then
                       GoTo ErrAgrDtosCtaTotGrab: lcError = oFactOrdenServicioPagos.MensajeError
                   End If
                   lnIdOrden888 = oDoFactOrdenServPagos.IdOrden
                   Do While Not oRsItemsCuenta.EOF And lnIdOrdenPago000 = oRsItemsCuenta!IdOrden
                        If lnNumeroDeItemFinal > Val(wxParametro102) Then
                                oDoFacturacionServicioPagos.IdOrdenPago = lnIdOrdenPago000
                                oDoFacturacionServicioPagos.idProducto = oRsItemsCuenta!idProducto
                                oDoFacturacionServicioPagos.IdUsuarioAuditoria = mo_DOComprobantePago.IdUsuarioAuditoria
                                If oFacturacionServicioPagos.EliminarXproducto(oDoFacturacionServicioPagos) = False Then
                                   GoTo ErrAgrDtosCtaTotGrab: lcError = oFacturacionServicioPagos.MensajeError
                                End If
                                If lnNroItemsXordenPago > Val(wxParametro102) Then
                                     If oFactOrdenServicioPagos.Insertar(oDoFactOrdenServPagos) = False Then
                                        GoTo ErrAgrDtosCtaTotGrab: lcError = oFactOrdenServicioPagos.MensajeError
                                     End If
                                     lnNroItemsXordenPago = 1
                                End If
                                oDoFacturacionServicioPagos.Cantidad = oRsItemsCuenta!Cantidad
                                oDoFacturacionServicioPagos.IdOrdenPago = oDoFactOrdenServPagos.IdOrdenPago
                                oDoFacturacionServicioPagos.idProducto = oRsItemsCuenta!idProducto
                                oDoFacturacionServicioPagos.IdUsuarioAuditoria = mo_DOComprobantePago.IdUsuarioAuditoria
                                oDoFacturacionServicioPagos.precio = oRsItemsCuenta!PrecioUnitario
                                oDoFacturacionServicioPagos.Total = oRsItemsCuenta!TotalPorPagar
                                If oFacturacionServicioPagos.Insertar(oDoFacturacionServicioPagos) = False Then
                                   GoTo ErrAgrDtosCtaTotGrab: lcError = oFacturacionServicioPagos.MensajeError
                                End If
                                lnIdOrdenPago111 = oDoFactOrdenServPagos.IdOrdenPago
                        Else
                                lnIdOrdenPago111 = oRsItemsCuenta!IdOrden
                        End If
                        lnImporteTotal111 = oRsItemsCuenta!TotalPorPagar
                        'actualiza EXONERACIONES, subTotal
                        lnImporteExonerado111 = 0
                        lnIdUsuarioExonera111 = 0
                        oRsExoneracionesTmp.Filter = "idOrden=" & lnIdOrden888 & _
                                                     " and idProducto=" & oRsItemsCuenta!idProducto & _
                                                     " and IdTipoFinanciamiento = 9 and IdEstadoFacturacion <> 9"
                        If oRsExoneracionesTmp.RecordCount > 0 Then
                           lnImporteExonerado111 = oRsExoneracionesTmp!TotalFinanciado
                           lnIdUsuarioExonera111 = oRsExoneracionesTmp!IdUsuarioAutoriza
                        End If
                        lbNuevoOrdenPago = True
                        If oRsOrdenPagoTmp.RecordCount > 0 Then
                           oRsOrdenPagoTmp.MoveFirst
                           oRsOrdenPagoTmp.Find "idOrdenPago=" & lnIdOrdenPago111
                           If Not oRsOrdenPagoTmp.EOF Then
                              lbNuevoOrdenPago = False
                           End If
                        End If
                        If lbNuevoOrdenPago = True Then
                           oRsOrdenPagoTmp.AddNew
                           oRsOrdenPagoTmp!IdOrdenPago = lnIdOrdenPago111
                           oRsOrdenPagoTmp!ImporteExonerado = lnImporteExonerado111
                           oRsOrdenPagoTmp!Subtotal = lnImporteTotal111
                           oRsOrdenPagoTmp!IdOrdenPagoAnterior = lnIdOrdenPago000
                        Else
                           oRsOrdenPagoTmp!ImporteExonerado = oRsOrdenPagoTmp!ImporteExonerado + lnImporteExonerado111
                           oRsOrdenPagoTmp!Subtotal = oRsOrdenPagoTmp!Subtotal + lnImporteTotal111
                        End If
                        If lnIdUsuarioExonera111 > 0 Then
                           oRsOrdenPagoTmp!idUsuarioExonera = lnIdUsuarioExonera111
                        End If
                        oRsOrdenPagoTmp.Update
                        '
                        oRsItemsCuenta.Fields!idordenPagoNEW = lnIdOrdenPago111
                        oRsItemsCuenta.Update
                        oRsItemsCuenta.MoveNext
                        lnNumeroDeItemFinal = lnNumeroDeItemFinal + 1
                        lnNroItemsXordenPago = lnNroItemsXordenPago + 1
                        If oRsItemsCuenta.EOF Then
                           Exit Do
                        End If
                   Loop
                   'actualiza exoneraciones, solo para aquellos que pasan del Máximo de Items
                   If (lnNumeroDeItemFinal - 1) > Val(wxParametro102) Then
                      oRsExoneracionesTmp.Filter = "idOrden=" & lnIdOrden888 & _
                                                   " and IdTipoFinanciamiento = 9 and IdEstadoFacturacion <> 9"
                      If oRsExoneracionesTmp.RecordCount > 0 Then
                         oRsOrdenPagoTmp.Filter = "IdOrdenPagoAnterior=" & lnIdOrdenPago000
                         If oRsOrdenPagoTmp.RecordCount > 0 Then
                            oRsOrdenPagoTmp.MoveFirst
                            Do While Not oRsOrdenPagoTmp.EOF
                               If mo_ReglasFacturacion.FactOrdenServicioPagosActualizaExoneracion(oRsOrdenPagoTmp!IdOrdenPago, _
                                                                         oRsOrdenPagoTmp!ImporteExonerado, _
                                                                         oRsOrdenPagoTmp!idUsuarioExonera, oConexion) = False Then
                                  GoTo ErrAgrDtosCtaTotGrab: lcError = mo_ReglasFacturacion.MensajeError
                               End If
                               oRsOrdenPagoTmp.MoveNext
                            Loop
                         End If
                      End If
                   End If
                   '
                Loop
                oRsItemsCuenta.Sort = ""
                oRsItemsCuenta.MoveFirst
                Do While Not oRsItemsCuenta.EOF
                   oRsItemsCuenta!IdOrden = oRsItemsCuenta!idordenPagoNEW
                   oRsItemsCuenta.Update
                   oRsItemsCuenta.MoveNext
                Loop
                'graba e imprime cada Boleta por cada RECETA, que ya no se parasan del MAXIMO DE ITEMS
                lnPagoACuenta = CCur(txtPagoACuenta.Text)
                oRsOrdenPagoTmp.Filter = ""
                oRsOrdenPagoTmp.MoveFirst
                Do While Not oRsOrdenPagoTmp.EOF
                    If (oRsOrdenPagoTmp!Subtotal - (oRsOrdenPagoTmp!ImporteExonerado + lnPagoACuenta)) <= 0 Then
                       lnAdelantos111 = oRsOrdenPagoTmp!Subtotal - oRsOrdenPagoTmp!ImporteExonerado
                       lnPagoACuenta = lnPagoACuenta - lnAdelantos111
                    Else
                       lnAdelantos111 = lnPagoACuenta
                       lnPagoACuenta = 0
                    End If
                    mo_DOComprobantePago.exoneraciones = oRsOrdenPagoTmp!ImporteExonerado
                    mo_DOComprobantePago.Adelantos = lnAdelantos111
                    mo_DOComprobantePago.Subtotal = oRsOrdenPagoTmp!Subtotal
                    mo_DOComprobantePago.Total = oRsOrdenPagoTmp!Subtotal - (oRsOrdenPagoTmp!ImporteExonerado + lnAdelantos111)
                    oRsItemsCuenta.Filter = "idOrdenPagoNew=" & oRsOrdenPagoTmp!IdOrdenPago
                    If oRsItemsCuenta.RecordCount > 0 Then
                            AgregarDatosPorCuentaTotalGrabaBoletaBS = oDllFactUCGestionCaja.CajaComprobantePagoCuentaAtencionAgregarS(mo_DOComprobantePago, _
                                                         mo_doCajaGestion, oRsItemsCuenta, _
                                                         lnIdCuentaAtencionAcerrar, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
                            If AgregarDatosPorCuentaTotalGrabaBoletaBS = True Then
                                ImpresionDelRecibo txtNroSerie, txtNroDocumento, sghServicio, lnParametrosImprimeBoleta, lbEsUnaFactura
                                mo_ReglasComunes.WaitSeconds Val(wxParametro501)
                                cmbIdTipoComprobante_Click
                                mo_DOComprobantePago.nrodocumento = UserControl.txtNroDocumento.Text
                                mo_DOComprobantePago.nroSerie = UserControl.txtNroSerie.Text
                            Else
                               lcError = lcError & " falta Boleta para ORDEN DE PAGO: " & Trim(Str(oRsOrdenPagoTmp!IdOrdenPago)) & Chr(13)
                            End If
                    End If
                    oRsOrdenPagoTmp.MoveNext
                Loop
            End If
            
            '******************************************************** imprime VARIAS boletas (fin)************
    Else
            '********* imprime UNA boleta - version anterior ************
            lbElNroItemsEsMenorAlMaximoDeBoleta = True
            If lbEsBoletaFarmacia = True Then
                AgregarDatosPorCuentaTotalGrabaBoletaBS = oDllFactUCGestionCaja.CajaComprobantePagoCuentaAtencionAgregarB(mo_DOComprobantePago, _
                                             oRsItemsCuenta, lnIdCuentaAtencionAcerrar, _
                                             mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
            Else
                AgregarDatosPorCuentaTotalGrabaBoletaBS = oDllFactUCGestionCaja.CajaComprobantePagoCuentaAtencionAgregarS(mo_DOComprobantePago, _
                                             mo_doCajaGestion, oRsItemsCuenta, _
                                             lnIdCuentaAtencionAcerrar, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
            End If
            mo_ReglasComunes.WaitSeconds Val(wxParametro501)
    End If
ErrAgrDtosCtaTotGrab:
    UserControl.Enabled = True
    UserControl.MousePointer = 1
    UserControl.KeyPreview = True
    
    If lcError <> "" Then
       MsgBox lcError, vbInformation, ""
    End If
    Set oFactOrdenesBienes = Nothing
    Set oDoFactOrdenesBienes = Nothing
    Set oFacturacionBienesPagos = Nothing
    Set oDoFacturacionBienesPagos = Nothing
    Set oFactOrdenServicioPagos = Nothing
    Set oDoFactOrdenServPagos = Nothing
    Set oFacturacionServicioPagos = Nothing
    Set oDoFacturacionServicioPagos = Nothing
    Set oRsExoneracionesTmp = Nothing
    Set oRsTmp2 = Nothing
    Set oRsOrdenPagoTmp = Nothing




End Function
Function Anulacion() As Boolean

    mo_DOComprobantePago.IdUsuarioAuditoria = ml_idUsuario
    If lbBoletaDeServicios Then
       Anulacion = mo_AdminCaja.CajaComprobantePagoServicioAnulaBoleta(mo_DOComprobantePago, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, 0)
       ms_MensajeError = mo_AdminCaja.MensajeError
    Else
       Anulacion = mo_ReglasFarmacia.CajaComprobantePagoBienesAnulaBoleta(mo_DOComprobantePago, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtRazonSocial.Text)
       ms_MensajeError = mo_ReglasFarmacia.MensajeError
    End If
    If mo_DOComprobantePago.idCuentaAtencion > 0 Then
       mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar mo_DOComprobantePago.idCuentaAtencion, False, 0
    End If
    'kike 2017
    If lbTieneLicenciaParaNotaCreditoYsunat = True Then
        Dim oExportar As New SIGHProxies.Procesos
        oExportar.ExportarFacturasBoletas "", "", mo_DOComprobantePago.nroSerie, mo_DOComprobantePago.nrodocumento, lbUsaResumenDiarioSunat
        Set oExportar = Nothing
    End If
    
    cmbIdTipoComprobante_Click
End Function

Function Devolucion() As Boolean
    
    Select Case ml_TipoProducto
    Case sghServicio
        Devolucion = mo_AdminCaja.CajaComprobantePagodevolucionOrdenServicio(mo_DOComprobantePago.IdComprobantePago, mo_DOComprobantePagoDevolucion, ml_idUsuario)
    Case sghbien
        Devolucion = mo_AdminCaja.CajaComprobantePagoDevolucionOrdenBienInsumo(mo_DOComprobantePago.IdComprobantePago, mo_DOComprobantePagoDevolucion, ml_idUsuario)
    End Select
    
End Function

Private Sub btnCancelar_Click()
    'Visible = False
End Sub

Private Sub Form_Load()
    
    Set mo_cmbFechaIngreso.MiComboBox = cmbFechaIngreso
    Set mo_cmbIdPuntoCarga.MiComboBox = cmbIdPuntoDeCarga
    Set mo_cmbIdTipoGenHistoriaClinica.MiComboBox = cmbIdTipoGenHistoriaClinica
    
    ConfigurarPuntosDeCarga
    ConfigurarTiposHistoriaClinica
    ConfigurarFechaIngreso
    
    mo_cmbIdPuntoCarga.BoundText = ml_PuntoCarga
    mo_cmbIdTipoGenHistoriaClinica.BoundText = 2

    ucFacturacionProductos.idUsuario = ml_idUsuario
    ucFacturacionProductos.Inicializar
    ucFacturacionProductos.idTipoFinanciamiento = ml_IdTipoFinanciamiento
    ucFacturacionProductos.TipoProducto = sghServicio
    ucFacturacionProductos.idPuntoCarga = ml_PuntoCarga
    
    CargarDatosAlFormulario
    
End Sub

Sub CargarDatosAlFormulario()
End Sub

Sub CargarDatosServiciosALosControlesPorIdOrden()
Dim oRsTmp1 As New Recordset

        If Trim(cmbOrdenes.Text) = "" Then
            MsgBox "Ingrese el numero de la orden que desea consultar", vbInformation, "Caja"
            Exit Sub
        End If
        
        ucFacturacionProductos.LimpiarGrilla
        
        'Carga datos de la orden
        Set oRsTmp1 = mo_AdminCaja.FactOrdenServicioPagosPendientesDePagoPorIdOrdenPago(Val(cmbOrdenes.Text))
        If oRsTmp1.RecordCount = 0 Then
            mb_ExistenDatos = False
            Exit Sub
        End If
        If wxParametro579 = "S" Then
           If oRsTmp1!idPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaAdmisionCE And oRsTmp1!FechaIngreso <> ldHoy Then
              MsgBox "No puede pagar una CITA hoy, tiene que hacerlo el " & oRsTmp1!FechaIngreso, vbInformation, ""
              Exit Sub
           End If
        End If
        
        lcDNIbuscado = IIf(IsNull(oRsTmp1!nrodocumento), "", oRsTmp1!nrodocumento) 'debb-05/12/2017
        mo_cmbIdPuntoCarga.BoundText = oRsTmp1.Fields!idPuntoCarga
        IdOrden = Val(cmbOrdenes.Text)
        ml_IdOrdenDespacho = oRsTmp1.Fields!IdOrden
        mb_ExistenDatos = True
        mo_DoAtencion.idAtencion = IIf(IsNull(oRsTmp1.Fields!idAtencion), 0, oRsTmp1.Fields!idAtencion)
        mo_DoAtencion.idPaciente = IIf(IsNull(oRsTmp1.Fields!idPaciente), 0, oRsTmp1.Fields!idPaciente)
        mo_DoAtencion.IdFormaPago = oRsTmp1.Fields!idTipoFinanciamiento
        Set mo_DoFactOrdenServPagos = mo_ReglasFacturacion.FactOrdenServicioPagosSeleccionarPorIdOrdenPago(oRsTmp1.Fields!IdOrdenPago)
        
           'Valida el estado de la orden
        Select Case mi_Opcion
        Case sghopcionespago.sghPagarOrdenExistente
            Select Case oRsTmp1.Fields!idestadofacturacion
            Case 1
            Case 4
                MsgBox "La orden ya ha sido PAGADA.", vbInformation, "Caja"
                Exit Sub
            Case 9
                MsgBox "La orden no puede ser pagada, se encuentra en estado ANULADO.", vbInformation, "Caja"
                Exit Sub
            End Select
        Case sghopcionespago.sghPagarCuentaExistente
        
        Case sghopcionespago.sghDevolucion
            Select Case oRsTmp1.Fields!idestadofacturacion
            Case 1
                MsgBox "La orden aun no ha sido PAGADA, solo se puede realizar devoluciones de ordenes PAGADAS.", vbInformation, "Caja"
                Exit Sub
            Case 4
                'Solo se pueden devolver ordenes pagadas
                'Verificar que tenga productos con autorizacion de devolucion
            Case 9
                MsgBox "No se puede realizar la devolución, la orden se encuentra en estado ANULADO.", vbInformation, "Caja"
                Exit Sub
            End Select
        Case sghopcionespago.sghAnulacion
            Select Case oRsTmp1.Fields!idestadofacturacion
            Case 1
                MsgBox "La orden aun no ha sido PAGADA, solo se puede anular de ordenes PAGADAS.", vbInformation, "Caja"
                Exit Sub
            Case 4
                'Solo se pueden devolver ordenes pagadas
                'Verificar que tenga productos con autorizacion de devolucion
            Case 9
                MsgBox "No se puede realizar la anulación, la orden ya se encuentra ANULADA.", vbInformation, "Caja"
                Exit Sub
            End Select
        End Select
         
        'Cargar datos del paciente y de la atencion
        ml_IdPaciente = IIf(IsNull(oRsTmp1.Fields!idPaciente), 0, oRsTmp1.Fields!idPaciente)
        txtNombres.Text = IIf(IsNull(oRsTmp1.Fields!ApellidoPaterno), "", oRsTmp1.Fields!ApellidoPaterno + " " + oRsTmp1.Fields!ApellidoMaterno + " " + oRsTmp1.Fields!PrimerNombre)
        txtNroHistoria.Text = IIf(IsNull(oRsTmp1.Fields!NroHistoriaClinica), 0, HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(oRsTmp1.Fields!NroHistoriaClinica)), False))
        txtRazonSocial = txtNombres
        txtNroCuenta.Text = IIf(IsNull(oRsTmp1.Fields!idCuentaAtencion), 0, oRsTmp1.Fields!idCuentaAtencion)
        ml_idCuentaAtencion = IIf(IsNull(oRsTmp1.Fields!idCuentaAtencion), 0, oRsTmp1.Fields!idCuentaAtencion)
        mo_cmbIdTipoFinanciamiento.BoundText = oRsTmp1.Fields!idTipoFinanciamiento
        If Not IsNull(oRsTmp1!EpsPorcentaje) Then
           txtObservaciones.Text = mo_AdminCaja.DevuelveMovimientoEnLabImaFar(Val(cmbOrdenes.Text), 1, oRsTmp1!idCuentaAtencion)
        End If
        'Cargar datos de los servicios
        Select Case mi_Opcion
        Case sghopcionespago.sghPagarOrdenExistente
        Case sghopcionespago.sghPagarCuentaExistente
        Case sghopcionespago.sghDevolucion
            ucFacturacionProductos.EstadosFacturacion = "5"    'Autorizados a devolver
            ucFacturacionProductos.TiposFinanciamiento = "1,5,9"
        Case sghopcionespago.sghAnulacion
            ucFacturacionProductos.EstadosFacturacion = "4"    'Pagados
            ucFacturacionProductos.TiposFinanciamiento = "1,5,9"
        End Select
        ucFacturacionProductos.IdOrdenPago = oRsTmp1.Fields!IdOrdenPago
        ucFacturacionProductos.CargaProductosPorIdOrdenPago
        txtExonerado.Text = mo_DoFactOrdenServPagos.ImporteExonerado
        If mi_Opcion <> sghPagarOrdenExistente Then    'debb-13/03/2012
           lnTotalGrid = md_Total + Val(txtExonerado.Text)
        End If
        ActualizaTotalApagar
        oRsTmp1.Close
        Set oRsTmp1 = Nothing
End Sub

Sub CargarDatosBienesALosControlesPorIdOrden()
Dim oDOFactOrdenBienInsumo As DoFactOrdenesBienes
Dim oDoPreVenta As New DoFarmPreVenta
Dim oPreVenta As New FarmPreVenta
Dim oConexion As New ADODB.Connection
Dim lcTipoVenta As String
        oConexion.Open sighEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        cmbOrdenes.Text = IIf(InStr(cmbOrdenes.Text, " ") > 0, Mid(cmbOrdenes.Text, 1, InStr(cmbOrdenes.Text, " ")), cmbOrdenes.Text)
        If Trim(cmbOrdenes.Text) = "" Then
            MsgBox "Ingrese el numero de la orden que desea consultar", vbInformation, "Caja"
            Exit Sub
        End If
        IdOrden = Val(cmbOrdenes.Text)
        txtTotal.Text = ""
        
        'Carga datos de la orden
        Select Case mi_Opcion
        Case sghopcionespago.sghPagarOrdenExistenteF
             UcFacturacionContado1.InHabilitaEdicionColumnasDelGrid = True
             UcFacturacionContado1.Inicializar
             UcFacturacionContado1.idPreVenta = 0
             UcFacturacionContado1.LimpiarGrilla
             'Set oDOFactOrdenBienInsumo = mo_ReglasFacturacion.FactOrdenesBienesInsumoSeleccionarPorIdPreVenta(Val(cmbOrdenes.Text))
              Set oDOFactOrdenBienInsumo = mo_ReglasFacturacion.FactOrdenesBienesInsumoSeleccionarPorIdPreVentaSoloUNO(Val(cmbOrdenes.Text))
        Case Else
             ucFacturacionProductos.LimpiarGrilla
             Set oDOFactOrdenBienInsumo = mo_ReglasFacturacion.FactOrdenesBienesInsumoSeleccionarPorId(Val(cmbOrdenes.Text))
        End Select
        If Not oDOFactOrdenBienInsumo Is Nothing Then
             ml_idCuentaAtencion = oDOFactOrdenBienInsumo.idCuentaAtencion
             ml_IdPaciente = oDOFactOrdenBienInsumo.idPaciente
            
'             ml_IdFormaPago = 1  'contado
'             ml_IdTipoFinanciamiento = ml_IdFormaPago
             Dim orstemp1 As New Recordset
             Set orstemp1 = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(ml_idCuentaAtencion, oConexion)
             If orstemp1.RecordCount > 0 Then
                lcTipoVenta = orstemp1.Fields!tipoventa
                If lcTipoVenta = "P" Then
                    ml_IdFormaPago = orstemp1.Fields!IdFormaPago
                Else
                    ml_IdFormaPago = 1  'contado
                End If
             Else
                ml_IdFormaPago = 1  'contado
             End If
             ml_IdTipoFinanciamiento = ml_IdFormaPago
             mo_cmbIdTipoFinanciamiento.BoundText = ml_IdTipoFinanciamiento
             
             oDoPreVenta.idPreVenta = oDOFactOrdenBienInsumo.idPreVenta
             ml_idPreVenta = oDOFactOrdenBienInsumo.idPreVenta
             Set oPreVenta.Conexion = oConexion
             If Not oPreVenta.SeleccionarPorId(oDoPreVenta) Then
                MsgBox "Problemas con tabla PRE-VENTA", vbInformation, "Caja"
                Exit Sub
             End If
             ml_IdFarmacia = oDoPreVenta.IdAlmacen
             Set mo_DOFactOrdenBienInsumo = oDOFactOrdenBienInsumo
             With mo_DOFactOrdenBienInsumo
                mo_cmbIdPuntoCarga.BoundText = mo_DOFactOrdenBienInsumo.idPuntoCarga
                 
                 mb_ExistenDatos = True
             End With
         Else
            mb_ExistenDatos = False
            Exit Sub
         End If
         
           'Valida el estado de la orden
        Select Case mi_Opcion
        Case sghopcionespago.sghPagarOrdenExistenteF
            Select Case mo_DOFactOrdenBienInsumo.idestadofacturacion
            Case 1
            Case 4
                MsgBox "La orden ya ha sido PAGADA.", vbInformation, "Caja"
                Exit Sub
            Case 9
                MsgBox "La orden no puede ser pagada, se encuentra en estado ANULADO.", vbInformation, "Caja"
                Exit Sub
            End Select
            If ml_IdFormaPago <> 1 Then
                If lcTipoVenta <> "P" Then
                    MsgBox "Solo se puede EMITIR BOLETA para DOCUMENTO=CONTADO", vbInformation, "Caja"

                    Exit Sub
                End If
            End If
        Case sghopcionespago.sghPagarCuentaExistente
        
        Case sghopcionespago.sghDevolucion
            Select Case mo_DOFactOrdenBienInsumo.idestadofacturacion
            Case 1
                MsgBox "La orden aun no ha sido PAGADA, solo se puede realizar devoluciones de ordenes PAGADAS.", vbInformation, "Caja"
                Exit Sub
            Case 4
                'Solo se pueden devolver ordenes pagadas
                'Verificar que tenga productos con autorizacion de devolucion
            Case 9
                MsgBox "No se puede realizar la devolución, la orden se encuentra en estado ANULADO.", vbInformation, "Caja"
                Exit Sub
            End Select
        
        Case sghopcionespago.sghAnulacion
            Select Case mo_DOFactOrdenBienInsumo.idestadofacturacion
            Case 1
                MsgBox "La orden aun no ha sido PAGADA, solo se puede realizar anulaciones de ordenes PAGADAS.", vbInformation, "Caja"
                Exit Sub
            Case 4
                'Solo se pueden devolver ordenes pagadas
                'Verificar que tenga productos con autorizacion de devolucion
            Case 9
                MsgBox "No se puede realizar la anulación, la orden se encuentra en estado ANULADO.", vbInformation, "Caja"
                Exit Sub
            End Select
            
        End Select
         
         
        'Cargar datos del paciente y de la atencion

        If ml_IdPaciente > 0 Then
'            Set mo_DOAtencion = mo_AdminAdmision.AtencionesSeleccionarPorId(mo_DOFactOrdenBienInsumo.idAtencion)
'            Set ucFacturacionProductos.Atencion = mo_DOAtencion
'            cmbFechaIngreso.Text = mo_DOAtencion.FechaIngreso
        End If
        txtNombres.Text = ""
        mo_cmbIdTipoGenHistoriaClinica.BoundText = ""
        txtNroHistoria.Text = ""
        'JR 1105 (2l)
        txtRazonSocial.Text = ms_Descripcion

        txtObservaciones.Text = ms_Descripcion
        If ml_IdPaciente > 0 Then
            Dim oDOCajaComprobantesPago As New DOCajaComprobantesPago
            Set mo_DOComprobantePago = mo_AdminCaja.ComprobantePagoSeleccionarPorId(mo_DOFactOrdenBienInsumo.IdComprobantePago, oConexion)
            Dim oDOPaciente As New doPaciente
            Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(ml_IdPaciente, oConexion)
            If Not oDOPaciente Is Nothing Then
                txtNombres.Text = oDOPaciente.ApellidoPaterno + " " + oDOPaciente.ApellidoMaterno + " " + oDOPaciente.PrimerNombre
                mo_cmbIdTipoGenHistoriaClinica.BoundText = oDOPaciente.idTipoNumeracion
                txtNroHistoria.Text = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(oDOPaciente.NroHistoriaClinica)), False)
                txtRazonSocial = txtNombres
            End If
        Else
        End If
        'Cargar datos de los BienInsumos
        Select Case mi_Opcion
        Case sghopcionespago.sghPagarOrdenExistenteF
            
            UcFacturacionContado1.Visible = True
            ucFacturacionProductos.Visible = False
            
            UcFacturacionContado1.IdAlmacen = ml_IdFarmacia
            UcFacturacionContado1.idPreVenta = oDOFactOrdenBienInsumo.idPreVenta
            'debb-18/05/2016 (inicio)
            UcFacturacionContado1.CargaProductosPorIdPreVenta lnTotalGrid
            lnTotalGrid = sighEntidades.DevuelveNumeroRedondeado(lnTotalGrid)
            'lnTotalGrid = oDoPreVenta.Total
            'debb-18/05/2016 (fin)
            txtExonerado.Text = oDOFactOrdenBienInsumo.ImporteExonerado
            ActualizaTotalApagar
            md_Total = CDbl(txtTotal.Text)
        Case sghopcionespago.sghPagarCuentaExistente
            UcFacturacionContado1.Visible = False
            ucFacturacionProductos.Visible = True
            ucFacturacionProductos.IdOrden = IdOrden
            ucFacturacionProductos.CargaProductosPorIdOrden
        Case sghopcionespago.sghDevolucion
            UcFacturacionContado1.Visible = False
            ucFacturacionProductos.Visible = True
            ucFacturacionProductos.EstadosFacturacion = "5"    'Autorizados a devolver
            ucFacturacionProductos.TiposFinanciamiento = "1,5,9"
            ucFacturacionProductos.IdOrden = IdOrden
            ucFacturacionProductos.CargaProductosPorIdOrden
        Case sghopcionespago.sghAnulacion
            UcFacturacionContado1.Visible = False
            ucFacturacionProductos.Visible = True
            'ucFacturacionProductos.EstadosFacturacion = "4"    'Pagados
            ucFacturacionProductos.EstadosFacturacion = IIf(mo_DOComprobantePago.IdTipoPago = 2, 6, 4)
            ucFacturacionProductos.TiposFinanciamiento = "1,2,3,5,9"
            ucFacturacionProductos.IdOrden = IdOrden
            ucFacturacionProductos.CargaProductosPorIdOrden
        End Select
        If mi_Opcion = sghPagarOrdenExistenteF And UserControl.txtRazonSocial.Text = "" Then
           UserControl.txtRazonSocial.Text = oDoPreVenta.Paciente
           UserControl.txtDni.Text = oDoPreVenta.DNI
        End If
        Set oDoPreVenta = Nothing
        Set oPreVenta = Nothing
        Set oConexion = Nothing
End Sub

'debb-18/02/2011
Private Sub txtNroHistoria_LostFocus()
    If txtNroHistoria.Text <> "" Then
       If Len(txtNroHistoria.Text) > 9 Or mo_Teclado.TextoEsSoloNumeros(txtNroHistoria.Text) = False Then
          MsgBox "La longitud no puede pasar de 9", vbInformation, "caja"
          txtNroHistoria.Text = ""
          Exit Sub
       End If

       txtNroCuenta.Text = ""
       LimpiarFormulario
       If cmbIdTipoGenHistoriaClinica.Text <> "" Then
            Dim oRsTmp As New Recordset
            Dim lnIdPacienteHallado As Long
            Dim oConexion As New Connection
            oConexion.Open sighEntidades.CadenaConexion
            oConexion.CursorLocation = adUseClient
            
            Set oRsTmp = mo_AdminAdmision.PacientesXnroHistoriaTipoNumeracion(Val(HCigualDNI_AgregaNUEVEaLaHistoria(txtNroHistoria.Text)), Val(mo_cmbIdTipoGenHistoriaClinica.BoundText), oConexion)
            If oRsTmp.RecordCount > 0 Then
               lnIdPacienteHallado = oRsTmp.Fields!idPaciente
               If Not IsNull(oRsTmp.Fields!nrodocumento) And Not IsNull(oRsTmp.Fields!IdDocIdentidad) Then
                  If oRsTmp.Fields!IdDocIdentidad = 1 Then

                     txtDni.Text = oRsTmp.Fields!nrodocumento
                  End If
               End If
               oRsTmp.Close
               Set oRsTmp = mo_ReglasFarmacia.FacturacionCuentasAtencionSeleccionarPorIdPaciente(lnIdPacienteHallado, oConexion, True)
               If oRsTmp.RecordCount > 0 Then
                  txtNroCuenta.Text = oRsTmp.Fields!idCuentaAtencion
               Else
                  MsgBox "Ese Nro de HISTORIA existe, pero no tiene Nro Cuenta", vbInformation, "Caja"
               End If
               oRsTmp.Close
               Set oRsTmp = Nothing
               txtNroCuenta_KeyPress 13
            Else
               MsgBox "Ese Nro de HISTORIA no existe", vbInformation, "Caja"
               oRsTmp.Close
               Set oRsTmp = Nothing
            End If
            oConexion.Close
            Set oConexion = Nothing
        Else
            MsgBox "Elija 'Tipo Historia'", vbInformation, "Caja"
        End If
    End If
End Sub

Sub LeerServiciosPorTipoDePago()

     Select Case mi_Opcion
        Case sghopcionespago.sghNuevoPagoConHistoria
        
        Case sghopcionespago.sghNuevoPagoSinHistoria
        
        Case sghopcionespago.sghPagarOrdenExistente
           CargarDatosServiciosALosControlesPorIdOrden
            txtRazonSocial.SetFocus
        Case sghopcionespago.sghPagarCuentaExistente
        
        Case sghopcionespago.sghDevolucion
            
            CargarDatosALosControlesPorNroSerieBoleta
        Case sghopcionespago.sghAnulacion
            
            CargarDatosALosControlesPorNroSerieBoleta
'             btnAceptar.SetFocus
        Case sghopcionespago.sghReimprimirComprobante
            CargarDatosALosControlesPorNroSerieBoleta
    End Select

End Sub

Sub LeerBienesPorTipoDePago()

     Select Case mi_Opcion
        Case sghopcionespago.sghNuevoPagoConHistoria
        
        Case sghopcionespago.sghNuevoPagoSinHistoria
        
        Case sghopcionespago.sghPagarOrdenExistenteF
           CargarDatosBienesALosControlesPorIdOrden
           txtRazonSocial.SetFocus
        Case sghopcionespago.sghPagarCuentaExistente
        
        Case sghopcionespago.sghDevolucion
            
            CargarDatosALosControlesPorNroSerieBoleta
        Case sghopcionespago.sghAnulacion
            
            CargarDatosALosControlesPorNroSerieBoleta
            btnAceptar.SetFocus
        Case sghopcionespago.sghReimprimirComprobante
            CargarDatosALosControlesPorNroSerieBoleta
    End Select

End Sub
Function BoletaTieneRegistradoLaboratorioImagenes(lnIdComprobantePago As Long) As Boolean
    Dim oRsTmp As New Recordset, lcSql As String
    BoletaTieneRegistradoLaboratorioImagenes = False
    Set oRsTmp = mo_AdminCaja.CajaComprobantesPagoXimagenes(lnIdComprobantePago)
    If oRsTmp.RecordCount > 0 And IsNull(oRsTmp.Fields!idCuentaAtencion) Then
       MsgBox "El Documento " & txtNserieB.Text & " - " & txtNdocumentoB.Text & " tiene registrado Movimiento en Imágenes" & Chr(13) & Chr(13) & "Fecha: " & oRsTmp.Fields!fecha & ",      N° Movimiento: " & oRsTmp.Fields!IdMovimiento, vbInformation, "Caja"
       BoletaTieneRegistradoLaboratorioImagenes = True
    Else
       oRsTmp.Close
       Set oRsTmp = mo_AdminCaja.CajaComprobantesPagoXlaboratorio(lnIdComprobantePago)
       If oRsTmp.RecordCount > 0 And IsNull(oRsTmp.Fields!idCuentaAtencion) Then
           MsgBox "El Documento " & txtNserieB.Text & " - " & txtNdocumentoB.Text & " tiene registrado Movimiento en Laboratorio" & Chr(13) & Chr(13) & "Fecha: " & oRsTmp.Fields!fecha & ",      N° Movimiento: " & oRsTmp.Fields!IdMovimiento, vbInformation, "Caja"
           BoletaTieneRegistradoLaboratorioImagenes = True
       End If
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
End Function
'***************daniel barrantes**************
'***************Cargar datos de BOLETA
'***************es usado para DEVOLUCION y ANULACION
Sub CargarDatosALosControlesPorNroSerieBoleta()
    Dim rsBuscaBoleta As Recordset
    Dim oReglasCaja As New SIGHNegocios.ReglasCaja
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion
    
    If Trim(txtNserieB.Text) = "" Or Trim(txtNdocumentoB.Text) = "" Then
       oConexion.Close
       Set oReglasCaja = Nothing
       Set oConexion = Nothing
       Exit Sub
    End If
    Set rsBuscaBoleta = oReglasCaja.CajaComprobantesPagoSeleccionarPorNroSerieNroDocumento(txtNserieB.Text, txtNdocumentoB.Text)
    If rsBuscaBoleta.RecordCount = 0 Then
       MsgBox "El Documento " & txtNserieB.Text & " - " & txtNdocumentoB.Text & " NO EXISTE", vbInformation, "Caja"
       oConexion.Close
       Set rsBuscaBoleta = Nothing
       Set oConexion = Nothing
       Exit Sub
    End If
    If BoletaTieneRegistradoLaboratorioImagenes(rsBuscaBoleta.Fields!IdComprobantePago) = True Then
       oConexion.Close
       Set rsBuscaBoleta = Nothing
       Set oConexion = Nothing
       Exit Sub
    End If
    txtNroSerie.Text = Trim(txtNserieB.Text)
    txtNroDocumento.Text = Trim(txtNdocumentoB.Text)
    lbBoletaDeServicios = IIf(rsBuscaBoleta.Fields!IdTipoOrden = 1, True, False)
    If lbBoletaDeServicios Then
       If wxParametro558 = "S" Then
            MsgBox "No se puede ANULAR SERVICIOS, debe usar NOTA DE CREDITO", vbInformation, ""
            oConexion.Close
            Set rsBuscaBoleta = Nothing
            Set oConexion = Nothing
            Exit Sub
       End If
       Set rsBuscaBoleta = oReglasCaja.CajaComprobantePagoServiciosPorNroSerieNroDocumento(txtNroSerie.Text, txtNroDocumento.Text)
    Else
       If wxParametro557 = "S" Then
            MsgBox "No se puede ANULAR FARMACIA, debe usar NOTA DE CREDITO", vbInformation, ""
            oConexion.Close
            Set rsBuscaBoleta = Nothing
            Set oConexion = Nothing
            Exit Sub
       End If
       Set rsBuscaBoleta = oReglasCaja.CajaComprobantePagoProductosPorNroSerieNroDocumento(txtNroSerie.Text, txtNroDocumento.Text)
    End If
    If rsBuscaBoleta.RecordCount = 0 Then
       MsgBox "El Documento " & txtNroSerie.Text & " - " & txtNroDocumento.Text & " NO EXISTE", vbInformation, "Caja"
       oConexion.Close
       Set oReglasCaja = Nothing
       Set rsBuscaBoleta = Nothing
       Set oConexion = Nothing
       Exit Sub
    ElseIf rsBuscaBoleta.Fields!idEstadoComprobante = 9 Then
       MsgBox "El Documento " & txtNroSerie.Text & " - " & txtNroDocumento.Text & " YA ESTA ANULADO", vbInformation, "Caja"
       oConexion.Close
       Set oReglasCaja = Nothing
       Set rsBuscaBoleta = Nothing
       Set oConexion = Nothing
       Exit Sub
    ElseIf rsBuscaBoleta.Fields!idEstadoComprobante = 1 Then
       MsgBox "La orden aun no ha sido PAGADA, solo se puede realizar anulaciones de ordenes PAGADAS.", vbInformation, "Caja"
       oConexion.Close
       Set oReglasCaja = Nothing
       Set rsBuscaBoleta = Nothing
       Set oConexion = Nothing
       Exit Sub
    ElseIf TienePaqueteDespachado(rsBuscaBoleta.Fields!IdComprobantePago) = True Then
       MsgBox "No se podrá Anular porque ya se ha Atendido Parte/Total del PAQUETE", vbInformation, "Caja"
       oConexion.Close
       Set oReglasCaja = Nothing
       Set rsBuscaBoleta = Nothing
       Set oConexion = Nothing
       Exit Sub
    ElseIf Not IsNull(rsBuscaBoleta.Fields!FechaCobranza) Then
       If Not (CDate(Format(rsBuscaBoleta.Fields!FechaCobranza, sighEntidades.DevuelveFechaSoloFormato_DMY)) = CDate(txtFechaBoleta.Text) And Val(mo_cmbIdTurno.BoundText) = rsBuscaBoleta.Fields!IdTurno) Then
          If Not (CDate(Format(rsBuscaBoleta.Fields!FechaCobranza + 0.6, sighEntidades.DevuelveFechaSoloFormato_DMY_HMS)) = CDate(lcBuscaParametro.RetornaFechaServidorSQLserver)) Then
                MsgBox "Sólo se puede realizar anulaciones en la misma Fecha y Turno de Emisión.", vbInformation, "Caja"
                If lbTrabajaComoCajero = True Then
                   txtFechaBoleta.Enabled = True
                End If
                oConexion.Close
                Set oReglasCaja = Nothing
                Set rsBuscaBoleta = Nothing
                Set oConexion = Nothing
                Exit Sub
          End If
       End If
    End If
    'Carga Cabecera
    LimpiarOpciones
    If IsNull(rsBuscaBoleta.Fields!NroHistoriaClinica) Then
       txtNroHistoria.Text = ""
    Else
       txtNroHistoria.Text = Trim(IIf(IsNull(rsBuscaBoleta.Fields!NroHistoriaClinica), "", HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rsBuscaBoleta.Fields!NroHistoriaClinica)), False)))
    End If
    
    ml_IdPaciente = IIf(IsNull(rsBuscaBoleta.Fields!idPaciente), 0, rsBuscaBoleta.Fields!idPaciente)
    txtRazonSocial.Text = Trim(IIf(IsNull(rsBuscaBoleta.Fields!razonSocial), "", rsBuscaBoleta.Fields!razonSocial))
    mo_DOComprobantePago.IdComprobantePago = rsBuscaBoleta.Fields!IdComprobantePago
    Set mo_DOComprobantePago = mo_AdminCaja.ComprobantePagoSeleccionarPorId(rsBuscaBoleta.Fields!IdComprobantePago, oConexion)

    'Carga Productos
    ucFacturacionProductos.PermiteAgregarItems = False
    ucFacturacionProductos.LimpiarGrilla
    If lbBoletaDeServicios Then
       ucFacturacionProductos.CargarItemsALaGrillaS rsBuscaBoleta
    Else
       ucFacturacionProductos.CargarItemsALaGrillaB rsBuscaBoleta
    End If
    If mi_Opcion = sghopcionespago.sghDevolucion Then
         ucFacturacionProductos.ActualizaDevolucionAutorizada rsBuscaBoleta
         Dim lnTotal As Double
         rsBuscaBoleta.MoveFirst
         Do While Not rsBuscaBoleta.EOF
            If Not IsNull(rsBuscaBoleta.Fields!cantidadDev) Then
            lnTotal = lnTotal + rsBuscaBoleta.Fields!PrecioUnitario * rsBuscaBoleta.Fields!cantidadDev
            End If
            rsBuscaBoleta.MoveNext
         Loop
         txtPagoACuenta.Text = 0
         txtTotal.Text = Format(lnTotal, "#######.#0")
         txtVuelto.Text = Format("0", "#######.#0")
    Else
          
         rsBuscaBoleta.MoveFirst
         txtPagoACuenta.Text = Format(rsBuscaBoleta.Fields!Adelantos, "#######.#0")
         txtExonerado.Text = Format(rsBuscaBoleta.Fields!exoneraciones, "#######.#0")
         txtTotal.Text = Format(rsBuscaBoleta.Fields!TotalBoleta, "#######.#0")
         txtVuelto.Text = Format(rsBuscaBoleta.Fields!TotalBoleta, "#######.#0")
    End If
    oConexion.Close
    Set oReglasCaja = Nothing
    Set rsBuscaBoleta = Nothing
    Set oConexion = Nothing
End Sub

Function TienePaqueteDespachado(lnIdComprobantePago As Long) As Boolean
    Dim oRsTmp As New Recordset
    TienePaqueteDespachado = False
    Set oRsTmp = mo_ReglasFacturacion.FacturacionPaquetesXidComprobantePago(lnIdComprobantePago)
    If oRsTmp.RecordCount > 0 Then
       TienePaqueteDespachado = True
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
    
End Function

Sub ConfigurarFechaIngreso()
    
    mo_cmbFechaIngreso.ListField = "DescripcionLarga"
    mo_cmbFechaIngreso.BoundColumn = "IdCuentaAtencion"

End Sub

Sub ConfigurarPuntosDeCarga()
    
    mo_cmbIdPuntoCarga.ListField = "Descripcion"
    mo_cmbIdPuntoCarga.BoundColumn = "IdPuntoCarga"
    Set mo_cmbIdPuntoCarga.RowSource = mo_ReglasComunes.SeleccionarPuntosDeCarga()

End Sub

Sub ConfigurarTiposHistoriaClinica()
        
        mo_cmbIdTipoGenHistoriaClinica.BoundColumn = "IdTipoNumeracion"
        mo_cmbIdTipoGenHistoriaClinica.ListField = "DescripcionLarga"
        Set mo_cmbIdTipoGenHistoriaClinica.RowSource = mo_AdminArchivoClinico.TiposGeneracionHistoriasSeleccionarTodos()

End Sub

Sub ConfigurarTipoComprobante()

    mo_cmbIdTipoComprobante.BoundColumn = "IdTipoComprobante"
    mo_cmbIdTipoComprobante.ListField = "Descripcion"
    Set mo_cmbIdTipoComprobante.RowSource = mo_AdminCaja.TiposComprobanteSeleccionarTodos()

End Sub

Sub CargarDatosServiciosALosControlesPorIdCuentaAtencion(lIdCuentaAtencion As Long)
Dim oDOCuentaAtencion As DOCuentaAtencion
Dim lcSql As String
Dim oRsTmp As New Recordset
Dim oConexion As New Connection
        oConexion.Open sighEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        
        ucFactServiciosPorCuenta.LimpiarGrilla
        ucFactBienesPorCuenta.LimpiarGrilla
        txtTotal.Text = ""
        
        'Carga datos de la orden
        Set oDOCuentaAtencion = mo_ReglasFacturacion.CuentasAtencionSeleccionarPorId(lIdCuentaAtencion, oConexion)
        
        If Not oDOCuentaAtencion Is Nothing Then
            Set mo_DOCuentaAtencion = oDOCuentaAtencion
             With mo_DOCuentaAtencion
                 mb_ExistenDatos = True
             End With
         Else
            mb_ExistenDatos = False
            Exit Sub
         End If
         
           'Valida el estado de la orden
        Select Case mi_Opcion
        Case sghopcionespago.sghPagarCuentaExistente
            Select Case mo_DOCuentaAtencion.idEstado
            Case 1
            Case 4
                MsgBox "La cuenta de atención se encuentra en estado PAGADO.", vbInformation, "Caja"
                Exit Sub
            Case 5
                MsgBox "La cuenta de atención se encuentra en estado CERRADO.", vbInformation, "Caja"
                Exit Sub
            Case 9
                MsgBox "La cuenta de atención se encuentra en estado ANULADO.", vbInformation, "Caja"
                Exit Sub
            End Select
        End Select
         
        cmbFechaIngreso.Text = mo_DOCuentaAtencion.FechaApertura
        
        ml_IdPaciente = mo_DOCuentaAtencion.idPaciente
        Set mo_DoPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(mo_DOCuentaAtencion.idPaciente, oConexion)
        If Not mo_DoPaciente Is Nothing Then
            txtNombres.Text = mo_DoPaciente.ApellidoPaterno + " " + mo_DoPaciente.ApellidoMaterno + " " + mo_DoPaciente.PrimerNombre
            mo_cmbIdTipoGenHistoriaClinica.BoundText = mo_DoPaciente.idTipoNumeracion
            txtNroHistoria.Text = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(mo_DoPaciente.NroHistoriaClinica)), False)
            txtRazonSocial = txtNombres
        End If
        txtTotal.Text = ""
        If lbCargaEstadoDeCuentaFarmacia = True Then
            'Solo CUENTA FARMACIA en "CAJA DE SERVICIOS"
            '*****opcion: "Pagar cuenta total (FARMACIA)"*****
            ucFactBienesPorCuenta.EstadosFacturacion = "1,3"    'Registrados y pendientes de pago
            ucFactBienesPorCuenta.TiposFinanciamiento = "1,2,3,5,9"
            ucFactBienesPorCuenta.idCuentaAtencion = lIdCuentaAtencion
            ucFactBienesPorCuenta.TipoProducto = sghbien
            ucFactBienesPorCuenta.CargaProductosPorIdCuentaAtencion
            txtExonerado.Text = ucFactBienesPorCuenta.DevuelveTotalImporteExonerado()
            tabFactProductosPorCuenta.TabVisible(1) = True
            tabFactProductosPorCuenta.Tab = 1
        Else
            'debb-17/02/2011
            If optServicios.Value = True And lbCargaEstadoDeCuentaFS = False Then
                'Solo CUENTA FARMACIA en "CAJA DE FARMACIA", 'Solo CUENTA SERVICIO en "CAJA DE SERVICIOS"
                '*****opcion: "Pagar cuenta total (SERVICIOS)"*****
                ucFactServiciosPorCuenta.EstadosFacturacion = "1,3"    'Registrados y pendientes de pago y pagos a cuenta
                ucFactServiciosPorCuenta.TiposFinanciamiento = "1,2,3,5,9"
                ucFactServiciosPorCuenta.idCuentaAtencion = lIdCuentaAtencion
                ucFactServiciosPorCuenta.TipoProducto = sghServicio
                ucFactServiciosPorCuenta.CargaProductosPorIdCuentaAtencion
                txtExonerado.Text = ucFactServiciosPorCuenta.DevuelveTotalImporteExonerado()
                tabFactProductosPorCuenta.TabVisible(0) = True
                tabFactProductosPorCuenta.Tab = 0
            Else
                '*******opcion: "Pagar cuenta total (Serv/Farm)"******
                'Solo CUENTA FARMACIA en "CAJA DE SERVICIOS"
                ucFactBienesPorCuenta.EstadosFacturacion = "1,3"    'Registrados y pendientes de pago
                ucFactBienesPorCuenta.TiposFinanciamiento = "1,2,3,5,9"
                ucFactBienesPorCuenta.idCuentaAtencion = lIdCuentaAtencion
                ucFactBienesPorCuenta.TipoProducto = sghbien
                ucFactBienesPorCuenta.CargaProductosPorIdCuentaAtencion 'Carga productos consumido en farmacia
                txtCtaFarmExonerado.Text = ucFactBienesPorCuenta.DevuelveTotalImporteExonerado()
                tabFactProductosPorCuenta.TabVisible(1) = True
                tabFactProductosPorCuenta.Tab = 1
                txtCtaFarmTfarmacia.Text = ucFactBienesPorCuenta.DevuelveTotalFS
                'Solo CUENTA SERVICIO en "CAJA DE SERVICIOS"
                ucFactServiciosPorCuenta.EstadosFacturacion = "1,3"    'Registrados y pendientes de pago y pagos a cuenta
                ucFactServiciosPorCuenta.TiposFinanciamiento = "1,2,3,5,9"
                ucFactServiciosPorCuenta.idCuentaAtencion = lIdCuentaAtencion
                ucFactServiciosPorCuenta.TipoProducto = sghServicio
                ucFactServiciosPorCuenta.CargaProductosPorIdCuentaAtencion 'Carga servicios consumidos
                txtCtaServExonerado.Text = ucFactServiciosPorCuenta.DevuelveTotalImporteExonerado()
                tabFactProductosPorCuenta.TabVisible(0) = True
                tabFactProductosPorCuenta.Tab = 0
                txtCtaServTservicio.Text = ucFactServiciosPorCuenta.DevuelveTotalFS
            End If
            'debb-25/02/2011
            txtTotal.Text = mo_ReglasFacturacion.RetornaTotalPagosPendientesPorNroCuentadebb(lIdCuentaAtencion, oConexion)
            txtEfectivo.Text = txtTotal.Text
            'debb-17/02/2011
        End If
        'Consulta el tipo de financiamiento por lIdCuentaAtencion de la tabla Tipofinanciamiento
        Set oRsTmp = mo_AdminAdmision.atencionesXtipoFinanciamiento(lIdCuentaAtencion, oConexion)
        If oRsTmp.RecordCount > 0 Then
           mo_cmbIdTipoFinanciamiento.BoundText = oRsTmp.Fields!IdFormaPago
           lblCuentaConSeguro.Caption = IIf(oRsTmp.Fields!generaPago = 1, "", "Con Seguro")
        End If
        oRsTmp.Close
        'debb-12/12/12
        '
        'Cuenta, solo para Paciente Pagante
        If lbCargaEstadoDeCuentaFS = False And lblCuentaConSeguro.Caption = "" Then
           CargaDatosDeTotalesDeLaCuenta lIdCuentaAtencion
        End If
        '
        Set oRsTmp = Nothing
        oConexion.Close
        Set oConexion = Nothing
        '
        txtTotal.Text = mo_Teclado.DevuelveImporteRedondeado(Val(txtTotal.Text), 2)
End Sub

'debb-17/02/2011
Sub CargaDatosDeTotalesDeLaCuenta(lIdCuentaAtencion As Long)
    Dim lnTotalDctosPorAdelantos As Double
    Dim lnTotalPagarFarmacia As Double
    Dim lnPagosXdevoluciones As Double
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion

    oConexion.CursorLocation = adUseClient
    'Cargar a lnTotalDctosPorAdelantos el total de pagos a cuenta(adelantos)
    lnTotalDctosPorAdelantos = mo_AdminCaja.RetornaTotalDescuentosPorAdelantosSegunCuenta(lIdCuentaAtencion, oConexion)
    lnPagosXdevoluciones = mo_ReglasFacturacion.RetornaImporteDePagosXdevolucionesPorNroCuenta(ml_idCuentaAtencion, oConexion)
    lnTotalDctosPorAdelantos = lnTotalDctosPorAdelantos - lnPagosXdevoluciones
    If (lnTotalDctosPorAdelantos) > 0 Then 'Si se registraron adelantos para la cuenta
    'Consulta el monto total de lo consumido en farmacia
       lnTotalPagarFarmacia = mo_ReglasFacturacion.RetornaTotalPagosFarmaciaPendientesPorNroCuentadebb(lIdCuentaAtencion)
       If tabFactProductosPorCuenta.Tab = 0 Then
          '****SERVICIOS: disminuir consumo de FARMACIA del Adelanto
          txtPagoACuenta.Text = lnTotalDctosPorAdelantos - lnTotalPagarFarmacia
       Else 'Cuando tabFactProductosPorCuenta.Tab = 1 -> es consumo en farmacia
          '****FARMACIA:
          lnTotalGrid = lnTotalPagarFarmacia 'lnTotalPagarFarmacia=135.12
          If (lnTotalDctosPorAdelantos) > lnTotalPagarFarmacia Then 'lnTotalDctosPorAdelantos=500
            'Si el el pago adelantado es mayor a lo consumido en farmacia, txtPagoACuenta será el total a lo consumido en farmacia
             txtPagoACuenta.Text = lnTotalPagarFarmacia 'txtPagoACuenta.Text= 135.12
          Else
          'Si el pago adelantado es menor (Ejm: 100) que lo consumido en farmacia (135.12)
          'entonces el txtPagoACuenta = al pago por adelanto --> txtPagoACuenta.Text=100
             txtPagoACuenta.Text = lnTotalDctosPorAdelantos
          End If
          'Aqui se declara la variable que tendra el pagoAdelantadoFarmacia= txtPagoACuenta.Text =135.12
          'Comrpobar si se puede insertar aqui el procedimiento almacenado
       End If
    End If
    ActualizaTotalApagar
    If txtTotal.Text < 0 Then
       txtPagoACuenta.Text = CCur(txtPagoACuenta.Text) + CCur(txtTotal.Text)
       ActualizaTotalApagar
    End If


End Sub

Sub CargarDatosServiciosALosControlesPorIdCuentaAtencionSoloFarmaciaEnTabServicio(lIdCuentaAtencion As Long)
Dim oDOCuentaAtencion As DOCuentaAtencion
Dim oConexion As New Connection
        oConexion.Open sighEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        
        ucFactServiciosPorCuenta.LimpiarGrilla
        ucFactBienesPorCuenta.LimpiarGrilla
    
        'Carga datos de la orden
        Set oDOCuentaAtencion = mo_ReglasFacturacion.CuentasAtencionSeleccionarPorId(lIdCuentaAtencion, oConexion)
        
        If Not oDOCuentaAtencion Is Nothing Then
            Set mo_DOCuentaAtencion = oDOCuentaAtencion
             With mo_DOCuentaAtencion
                 mb_ExistenDatos = True
             End With
         Else
            mb_ExistenDatos = False
            Exit Sub
         End If
         
           'Valida el estado de la orden
        Select Case mi_Opcion
        Case sghopcionespago.sghPagarCuentaExistente
            Select Case mo_DOCuentaAtencion.idEstado
            Case 1
            Case 4
                MsgBox "La cuenta de atención se encuentra en estado PAGADO.", vbInformation, "Caja"
                Exit Sub
            Case 5
                MsgBox "La cuenta de atención se encuentra en estado CERRADO.", vbInformation, "Caja"
                Exit Sub
            Case 9
                MsgBox "La cuenta de atención se encuentra en estado ANULADO.", vbInformation, "Caja"
                Exit Sub
            End Select
        End Select
         
        cmbFechaIngreso.Text = mo_DOCuentaAtencion.FechaApertura
        
        
        Set mo_DoPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(mo_DOCuentaAtencion.idPaciente, oConexion)
        If Not mo_DoPaciente Is Nothing Then
            txtNombres.Text = mo_DoPaciente.ApellidoPaterno + " " + mo_DoPaciente.ApellidoMaterno + " " + mo_DoPaciente.PrimerNombre
            mo_cmbIdTipoGenHistoriaClinica.BoundText = mo_DoPaciente.idTipoNumeracion
            txtNroHistoria.Text = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(mo_DoPaciente.NroHistoriaClinica)), False)
            txtRazonSocial = txtNombres
        End If
         
        ucFactBienesPorCuenta.EstadosFacturacion = "1,3"    'Registrados y pendientes de pago
        ucFactBienesPorCuenta.TiposFinanciamiento = "1,2,3,5,9"
        ucFactBienesPorCuenta.idCuentaAtencion = lIdCuentaAtencion
        ucFactBienesPorCuenta.TipoProducto = sghbien
        ucFactBienesPorCuenta.CargaProductosPorIdCuentaAtencion
        tabFactProductosPorCuenta.TabVisible(1) = True
        tabFactProductosPorCuenta.Tab = 1
                

End Sub



Sub MuestraTabEmisionDocumentos(lbVisible As Boolean)
        UserControl.tabGestionCaja.TabVisible(1) = lbVisible
        If lbVisible = True Then 'Frank 24082015
            PermisosAccesoNotaCredito
        Else
            UserControl.tabGestionCaja.TabVisible(2) = False 'Frank 24082015
        End If
End Sub

'***************daniel barrantes**************
'***************Chequea si EXISTE BOLETA antes de GRABAR los datos en las tablas
'***************
Function ExisteComprobantePagoPorNroSerieDocumento() As Boolean
    Dim rsBuscaBoleta As Recordset
    Dim oReglasCaja As New SIGHNegocios.ReglasCaja
    Dim lbSigue As Boolean
    Dim lnLen As Integer
    lbSigue = False
    txtNroSerie.Text = Trim(txtNroSerie.Text)
    txtNroDocumento.Text = Trim(txtNroDocumento.Text)
    Do While lbSigue = False
        Set rsBuscaBoleta = oReglasCaja.CajaComprobantePagoPorSerieDocumento(txtNroSerie.Text, txtNroDocumento.Text)
        If rsBuscaBoleta.RecordCount = 0 Then
           ExisteComprobantePagoPorNroSerieDocumento = False
           lbSigue = True
        Else
           ExisteComprobantePagoPorNroSerieDocumento = True
           If MsgBox("El comprobante Nro " & Trim(txtNroSerie.Text) & " - " & Trim(txtNroDocumento.Text) & Chr(13) & " YA FUE EMITIDO ANTERIORMENTE" & Chr(13) & " !! cambie el NRO SERIE Y NRO DOCUMENTO" & Chr(13) & Chr(13) & "Desea Continuar con el Siguiente Número ? ", vbQuestion + vbYesNo, "CAJA") = vbYes Then
                lnLen = Len(txtNroDocumento.Text)
                txtNroDocumento.Text = Right("00000000" & Trim(Str(Val(txtNroDocumento.Text) + 1)), lnLen)
           Else
                lbSigue = True
           End If
        End If
        rsBuscaBoleta.Close
    Loop
End Function


Sub ActualizaTotalApagar()
    If txtExonerado.Text = "" Then
       txtExonerado.Text = "0"
    End If
    If txtPagoACuenta.Text = "" Then
       txtPagoACuenta.Text = "0"
    End If
    txtTotal.Text = lnTotalGrid - CCur(txtExonerado.Text) - CCur(txtPagoACuenta.Text)
    txtTotal.Text = sighEntidades.DevuelveNumeroRedondeado(CCur(txtTotal.Text))   'debb-mayo2014
    txtEfectivo.Text = txtTotal.Text
End Sub


Sub ConfiguraPermisos()
    'PERMISOS
    Dim oRsPermisos As New Recordset
    Set oRsPermisos = mo_ReglasSeguridad.UsuariosRolesSeleccionarPermisosTodos(ml_idUsuario)
    UserControl.tabGestionCaja.TabVisible(1) = False
    UserControl.tabGestionCaja.TabVisible(2) = False 'Frank 24082015
    lbTienePermisoSoloParaBoletaFarmacia = False
    lbTienePermisoReimprimeBoleta = False
    lbTienePermisoExonerarPacExterno = False
    lbTienePermisoSoloParaBoletaServicio = False
    optRealizarAnulacion.Enabled = False
    lbTrabajaComoCajero = False
    lbPuedeVerVistaPrevia = True
    If oRsPermisos.RecordCount > 0 Then
       Do While Not oRsPermisos.EOF
          Select Case oRsPermisos.Fields!IdPermiso
          Case 208    'Caja - NO puede ver VISTA PREVIA DE BOLETA
               lbPuedeVerVistaPrevia = False
          Case 206    'Caja - Ver TAB 'Registro de Comprobante'
               UserControl.tabGestionCaja.TabVisible(1) = True
          Case 207    'Caja - Ver TAB ''Devolución por Nota de Crédito Frank 24082015
               UserControl.tabGestionCaja.TabVisible(2) = True 'Frank 24082015
          Case 366    'Caja - Sólo Emite Boleta p' Pre-Venta Farmacia
               lbTienePermisoSoloParaBoletaFarmacia = True
          Case 367    'Caja - Reimprime Boleta
               lbTienePermisoReimprimeBoleta = True
          Case 368    'Caja - Pacientes Externos - permite ingresar exoneracion
               lbTienePermisoExonerarPacExterno = True
          Case 369    'Caja - Sólo Emite Boletas de Servicios - sólo CPT
               lbTienePermisoSoloParaBoletaServicio = True
          Case 205    'Caja - Realiza Anulaciones
               optRealizarAnulacion.Enabled = True
          Case 1000   'Cja - Trabaja Como Cajero
               lbTrabajaComoCajero = True
          End Select
          oRsPermisos.MoveNext
       Loop
    End If
    Set oRsPermisos = Nothing
End Sub

'***************** GalenHos v.3.0 (inicio)*****************
'*******impresion de factura **********
Sub CargaDatosAlObjetosDeDatosFactura()
    If lbEsUnaFactura = True Then
       Dim lnIGV1 As Double
       If mi_Opcion = sghPagarOrdenExistenteF Then
            '*********Farmacia**************
            If lbFacturaSinIGV = True Then
                mo_DOComprobantePago.Subtotal = mo_DOComprobantePago.Total
                mo_DOComprobantePago.IGV = 0
            Else
'                If wxParametro533 = "S" Then
'                    'aumenta el IGV al importe total
'                    mo_DOComprobantePago.Subtotal = mo_DOComprobantePago.Total
'                    mo_DOComprobantePago.IGV = Round(lnIGV * mo_DOComprobantePago.Total / 100, 2)
'                    mo_DOComprobantePago.Total = mo_DOComprobantePago.Subtotal + mo_DOComprobantePago.IGV
'                Else
                    'NO aumenta el IGV al importe total
                    lnIGV1 = lnIGV / 100
                    mo_DOComprobantePago.Subtotal = Round(mo_DOComprobantePago.Total / (lnIGV1 + 1), 2)
                    mo_DOComprobantePago.IGV = Round(mo_DOComprobantePago.Total * lnIGV1 / (lnIGV1 + 1), 2)
 '               End If
            End If
        Else
            '************Servicios*********
            If lbEstaCajaUsaDescripcionLarga = True Then
                'NO aumenta el IGV al importe total
                If lnMontoIGV99 > 0 Then
                    lnIGV1 = lnIGV / 100
                    mo_DOComprobantePago.Subtotal = mo_DOComprobantePago.Total - lnMontoIGV99
                    mo_DOComprobantePago.IGV = lnMontoIGV99
                End If
                mo_DOComprobantePago.TieneCredito = IIf(lbTieneCredito99 = True, "C", "")
            Else
                If lbFacturaSinIGV = True Then
                    'Servicios - No hay IGV en FACTURA DE SERVICIOS
                    mo_DOComprobantePago.Subtotal = mo_DOComprobantePago.Total
                    mo_DOComprobantePago.IGV = 0
                Else
                    If wxParametro533 = "S" Then
                        'aumenta el IGV al importe total
                        mo_DOComprobantePago.Subtotal = mo_DOComprobantePago.Total
                        mo_DOComprobantePago.IGV = Round(lnIGV * mo_DOComprobantePago.Total / 100, 2)
                        mo_DOComprobantePago.Total = mo_DOComprobantePago.Subtotal + mo_DOComprobantePago.IGV
                    Else
                        'NO aumenta el IGV al importe total
                        lnIGV1 = lnIGV / 100
                        mo_DOComprobantePago.Subtotal = Round(mo_DOComprobantePago.Total / (lnIGV1 + 1), 2)
                        mo_DOComprobantePago.IGV = Round(mo_DOComprobantePago.Total * lnIGV1 / (lnIGV1 + 1), 2)
                    End If
                End If
            End If
        End If
    Else
        mo_DOComprobantePago.ruc = ""
    End If
End Sub
'***************** GalenHos v.3.0 (fin)*****************

Private Sub UserControl_Show()
       If lbTienePermisoSoloParaBoletaFarmacia = True Then
               optNuevoOrdenPagoConHistoria.Enabled = False
               optCobrarOrdenExistente.Enabled = False
               optRealizarAnulacion.Enabled = False
               optPagarEstadoDeCTAFarmacia.Enabled = False
               optPagarEstadoDeCuenta.Enabled = False
               optNuevoOrdenPagoSinHistoria.Enabled = False
               optCobrarOrdenExistente.Value = True
      End If
End Sub

Private Sub chkGeneraPreventaServ_Click()
    If chkGeneraPreventaServ.Value = 1 Then
       GeneraPreventaServicio
    End If
End Sub

Sub GeneraPreventaServicio()
    Dim lnPreventaServicio As Long, lnIdServicio As Long
    Dim lcSql As String
    Dim oRsTmp As New Recordset
    '
    If ucFacturacionProductos.FacturacionProductos.RecordCount <= 0 Or CCur(txtTotal.Text) = 0 Then
       MsgBox "Debe ingresar al menos un PRODUCTO para la PreVenta", vbInformation, "Caja"
       Exit Sub
    End If
    'Busca el Servicio que corresponde al Primer Item de la Lista
    'dicho Item deberá tener un "Punto de Carga"
    'dicho Punto de Carga deberá tener un Id Servicio (Hosp/Emerg/Ce/otros)
    'Sino tiene deberá registrarlo en Fact-Config-->Catalogo de Servicio-->elegir Item-->Modificar
    ucFacturacionProductos.FacturacionProductos.MoveFirst
    Set oRsTmp = mo_ReglasComunes.FactCatalogoServiciosPtosXidProducto(ucFacturacionProductos.FacturacionProductos.Fields!idProducto)
    lnIdServicio = 0
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          If oRsTmp.Fields!IdServicio > 0 Then
             lnIdServicio = oRsTmp.Fields!IdServicio
             Exit Do
          End If
          oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
    If lnIdServicio = 0 Then
       MsgBox "El Producto: " & ucFacturacionProductos.FacturacionProductos.Fields!NombreProducto & " no tiene Punto de Carga" & Chr(13) & "Use la opcion: 'Fact_Config-->CatalogoServicios'  para registrarlo", vbInformation, "Caja"
       Exit Sub
    End If
    '
    If mo_AdminCaja.CajaGeneraPreventaServicio(lnPreventaServicio, lnIdServicio, Val(mo_cmbIdTipoFinanciamiento.BoundText), CCur(txtTotal.Text), ucFacturacionProductos.FacturacionProductos, ml_idUsuario, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc) Then
       MsgBox "Se generó la Preventa N° " & lnPreventaServicio
       optNuevoOrdenPagoSinHistoria_Click 1
    Else
       MsgBox mo_AdminCaja.MensajeError
    End If
End Sub

Sub ConfigurarTipoFinanciamiento()
       mo_cmbIdTipoFinanciamiento.BoundColumn = "IdTipoFinanciamiento"
       mo_cmbIdTipoFinanciamiento.ListField = "Descripcion"
       Set mo_cmbIdTipoFinanciamiento.RowSource = mo_ReglasFacturacion.TiposFinanciamientosSeleccionarPorGeneraPagos(sghTodosLosQuePaganEnCaja)
End Sub

'mgaray201513
Public Sub ActivarTabGestionCaja(lIndex As Long)
    tabGestionCaja.Tab = lIndex
End Sub

Public Sub ActivarOrdenExistenteFS()
    optOrdenExistenteFS.Value = True
End Sub

Public Function AsignarNroOrden(sNroOrden As String)
    cmbOrdenes.Text = sNroOrden
End Function

Public Sub BuscarOrdenExistente()
    cmbOrdenes_KeyPress 13
End Sub

'FRANK 23082015
Private Sub txtNotaSerie_KeyPress(KeyAscii As Integer)
   If Len(txtNotaSerie.Text) > 0 And KeyAscii = 13 And txtNotaDocumento.Text <> "" Then
      BuscarNotaCredito
'   ElseIf Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
'       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
'           KeyAscii = 0
'       End If
   End If
End Sub

Private Sub txtNotaDocumento_KeyPress(KeyAscii As Integer)
   If Len(txtNotaDocumento.Text) > 0 And KeyAscii = 13 Then
      BuscarNotaCredito
'   ElseIf Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
'       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
'           KeyAscii = 0
'       End If
   End If
End Sub

Sub BuscarNotaCredito()
    If Trim(txtNotaSerie.Text) = "" Then
        MsgBox "Ingrese el Nro de Serie de la Nota", vbInformation, lcTituloNotaCredito
        Exit Sub
    End If
    If Trim(txtNotaDocumento.Text) = "" Then
        MsgBox "Ingrese el Nro de Comprobante de la Nota", vbInformation, lcTituloNotaCredito
        Exit Sub
    End If
    MousePointer = 11
    RealizarBusquedaNota
    MousePointer = 1
End Sub

Sub RealizarBusquedaNota()
    Dim orsTemp As New Recordset
    Dim oDoCaja As New DOCajaCaja
    Set orsTemp = Nothing
    fraNota.Tag = ""
    LimpiarDatosNotasCredito
    Set orsTemp = mo_AdminCaja.NotaCreditoRegistrosTotalesPorNumero(Trim(txtNotaSerie.Text), Trim(txtNotaDocumento.Text))
    If orsTemp.RecordCount = 0 Then
        MsgBox "El Número de Serie y Documento Ingresado no existe, por favor ingrese otro.", vbInformation, lcTituloNotaCredito
        Exit Sub
    End If
    If orsTemp.RecordCount > 0 Then
        Do While Not orsTemp.EOF
            If orsTemp.Fields!IdEstadoNota = sghEstadoNotaCredito.anulado Then
                MsgBox "No puede cargar la nota de crédito " + Trim(txtNotaSerie.Text) + "-" + Trim(txtNotaDocumento.Text) + " porque fue anulada.", vbInformation, lcTituloNotaCredito
                Set orsTemp = Nothing
                Exit Sub
            End If
            chkRevertirPagoNota.Visible = IIf(orsTemp.Fields!IdEstadoNota = sghEstadoNotaCredito.Canjeado, True, False)
            If orsTemp.Fields!IdEstadoNota = sghEstadoNotaCredito.Canjeado Then
                If Not IsNull(orsTemp.Fields!IdCaja) Then
                    Set oDoCaja = mo_AdminCaja.CajaSeleccionarPorId(orsTemp.Fields!IdCaja)
                    MsgBox "La nota de crédito " + Trim(txtNotaSerie.Text) + "-" + Trim(txtNotaDocumento.Text) + " ya fue canjeada en la caja " & oDoCaja.descripcion & " el " & orsTemp.Fields!fechapagado, vbExclamation, lcTituloNotaCredito
                    If Not (orsTemp.Fields!IdCaja = mo_doCajaGestion.IdCaja And _
                            orsTemp.Fields!IdCajero = mo_doCajaGestion.IdCajero And _
                            orsTemp.Fields!IdTurno = mo_doCajaGestion.IdTurno And _
                            Format(orsTemp.Fields!fechapagado, sighEntidades.DevuelveFechaSoloFormato_DMY) = Format(mo_doCajaGestion.FechaApertura, sighEntidades.DevuelveFechaSoloFormato_DMY)) Then
                            chkRevertirPagoNota.Visible = False
                    End If
                    Set oDoCaja = Nothing
                End If
            End If
            fraNota.Tag = orsTemp.Fields!IdNota
            txtSerieNota.Text = Trim(txtNotaSerie.Text)
            txtDocumentoNota.Text = Trim(txtNotaDocumento.Text)
            txtEstadoNota.Text = IIf(IsNull(orsTemp.Fields!EstadoNota), "", orsTemp.Fields!EstadoNota)
            txtNotaRazonSocial.Text = IIf(IsNull(orsTemp.Fields!razonSocial), "", orsTemp.Fields!razonSocial)
            txtNotaCredDebDireccion.Text = IIf(IsNull(orsTemp.Fields!Direccion), "", orsTemp.Fields!Direccion)
            txtNotaRuc.Text = IIf(IsNull(orsTemp.Fields!ruc), "", orsTemp.Fields!ruc)
            txtNotaFechaEmision.Text = orsTemp.Fields!FechaAprueba
            txtNotaMotivo.Text = orsTemp.Fields!Motivo
            txtNotaConcepto.Text = IIf(IsNull(orsTemp.Fields!Observaciones), "", orsTemp.Fields!Observaciones)
            txtNotaTotal.Text = orsTemp.Fields!Total
            txtNotaCredAprueba.Text = orsTemp.Fields!EmpleadoAutoriza + " (DNI=" + IIf(IsNull(orsTemp.Fields!DNI), "", Trim(orsTemp.Fields!DNI)) + ")"
            
            Set mo_DoNotaCreditoDebito = mo_AdminCaja.NotaCreditoDebitoSeleccionarPorId(orsTemp.Fields!IdNota)
            If mo_DoNotaCreditoDebito.IdEstadoNota = sghEstadoNotaCredito.Canjeado Then
                If Not (mo_DoNotaCreditoDebito.IdCaja = mo_doCajaGestion.IdCaja And _
                        mo_DoNotaCreditoDebito.IdCajero = mo_doCajaGestion.IdCajero And _
                        mo_DoNotaCreditoDebito.IdTurno = mo_doCajaGestion.IdTurno And _
                        Format(mo_DoNotaCreditoDebito.fechapagado, sighEntidades.DevuelveFechaSoloFormato_DMY) = Format(mo_doCajaGestion.FechaApertura, sighEntidades.DevuelveFechaSoloFormato_DMY)) Then
                        chkRevertirPagoNota.Visible = False
                End If
            End If
                        
            orsTemp.MoveNext
        Loop
        txtNotaSerie.Text = ""
        txtNotaDocumento.Text = ""
    End If

End Sub

Private Sub btnAceptarPagoNota_Click()
    If lbLaDevolucionNCesAutomatica = False Then
        If btnAceptarPagoNota.Enabled = False Then
           Exit Sub
        End If
        If chkRevertirPagoNota.Visible = True And chkRevertirPagoNota.Value = 1 Then
            If MsgBox("Por favor confirmar, ¿Realmente desea grabar la reversión de la nota de crédito?", vbQuestion + vbYesNo, lcTituloNotaCredito) = vbNo Then
                Exit Sub
            End If
        Else
            If MsgBox("Por favor confirmar, ¿Realmente desea grabar la devolución de dinero por Nota de Crédito?", vbQuestion + vbYesNo, lcTituloNotaCredito) = vbNo Then
                Exit Sub
            End If
        End If
    End If
    If ValidarDatosObligNota() Then
        If ValidarReglasNota() Then
             CargaDatosAlObjDatosNotaCredito
             If GrabarPagoNotaCredito() Then
                 If lbLaDevolucionNCesAutomatica = False Then
                 MsgBox "Se Registró la devolución por Nota de Crédito" + Chr(13) + "La Nota de Crédito paso al estado Canjeado, se procede a devolver el efectivo.", vbInformation, lcTituloNotaCredito
                 End If
                 'kike 2017
                 If lbTieneLicenciaParaNotaCreditoYsunat = True Then
                    Dim oExportar As New SIGHProxies.Procesos
                    oExportar.ExportarNotasCredito "", "", mo_DoNotaCreditoDebito.nroSerie, mo_DoNotaCreditoDebito.nrodocumento, _
                                                   lbUsaResumenDiarioSunat
                    Set oExportar = Nothing
                 End If
                 
                 LimpiarDatosNotasCredito
             Else
                 MsgBox "No se pudo grabar los cambios a la Nota de Crédito " + Chr(13) + ms_MensajeError, vbExclamation, lcTituloNotaCredito
             End If
        End If
    End If
End Sub

Private Sub btnLimpiarNota_Click()
    LimpiarDatosNotasCredito
End Sub

Sub LimpiarDatosNotasCredito()
    txtSerieNota.Text = ""
    txtDocumentoNota.Text = ""
    fraNota.Tag = ""
    txtNotaRazonSocial.Text = ""
    txtNotaCredDebDireccion.Text = ""
    txtNotaRuc.Text = ""
    txtNotaFechaEmision.Text = ""
    txtNotaMotivo.Text = ""
    txtNotaConcepto.Text = ""
    txtNotaTotal.Text = ""
    chkRevertirPagoNota.Visible = False
    chkRevertirPagoNota.Value = 0
    txtEstadoNota.Text = ""
    txtNotaCredAprueba.Text = ""
End Sub

Sub CargaDatosAlObjDatosNotaCredito()
    If mo_DoNotaCreditoDebito.IdEstadoNota = sghEstadoNotaCredito.Aprobado Then
        mo_DoNotaCreditoDebito.fechapagado = lcBuscaParametro.RetornaFechaServidorSQL & " " & lcBuscaParametro.RetornaHoraServidorSQLserverFormatoGalenhos 'Actualizado FCV 30032015
        mo_DoNotaCreditoDebito.IdEstadoNota = sghEstadoNotaCredito.Canjeado
        If lbLaDevolucionNCesAutomatica = True Then
            mo_DoNotaCreditoDebito.IdCaja = AutomNC_lnIdCaja
            mo_DoNotaCreditoDebito.IdCajero = AutomNC_ml_IdUsuario
            mo_DoNotaCreditoDebito.IdGestionCaja = AutomNC_lnIdGestionCaja
            mo_DoNotaCreditoDebito.IdTurno = AutomNC_lnIdTurno
        Else
            mo_DoNotaCreditoDebito.IdCaja = Val(mo_cmbIdCaja.BoundText)
            mo_DoNotaCreditoDebito.IdCajero = ml_idUsuario
            mo_DoNotaCreditoDebito.IdGestionCaja = mo_doCajaGestion.IdGestionCaja
            mo_DoNotaCreditoDebito.IdTurno = Val(mo_cmbIdTurno.BoundText)
        End If
    Else
        If chkRevertirPagoNota.Visible = True And chkRevertirPagoNota.Value = 1 Then
            mo_DoNotaCreditoDebito.fechapagado = 0
            mo_DoNotaCreditoDebito.IdCaja = 0
            mo_DoNotaCreditoDebito.IdCajero = 0
            mo_DoNotaCreditoDebito.IdEstadoNota = sghEstadoNotaCredito.Aprobado
            mo_DoNotaCreditoDebito.IdGestionCaja = 0
            mo_DoNotaCreditoDebito.IdTurno = 0
        End If
    End If
    mo_DoNotaCreditoDebito.IdUsuarioAuditoria = 738 ' ml_IdUsuario
End Sub

Function GrabarPagoNotaCredito() As Boolean
    MousePointer = 11
    GrabarPagoNotaCredito = False
    GrabarPagoNotaCredito = mo_AdminCaja.NotaCreditoDebitoModificar(mo_DoNotaCreditoDebito)
    ms_MensajeError = mo_AdminCaja.MensajeError
    MousePointer = 1
End Function


Function ValidarReglasNota() As Boolean
    Dim sMensaje As String
    ValidarReglasNota = False
    Dim oMensaje As New SIGHNegocios.clMensaje
    If mo_DoNotaCreditoDebito.IdEstadoNota = sghEstadoNotaCredito.Canjeado Then
        If Not (mo_DoNotaCreditoDebito.IdCaja = mo_doCajaGestion.IdCaja And _
                mo_DoNotaCreditoDebito.IdCajero = mo_doCajaGestion.IdCajero And _
                mo_DoNotaCreditoDebito.IdTurno = mo_doCajaGestion.IdTurno And _
                Format(mo_DoNotaCreditoDebito.fechapagado, sighEntidades.DevuelveFechaSoloFormato_DMY) = Format(mo_doCajaGestion.FechaApertura, sighEntidades.DevuelveFechaSoloFormato_DMY)) Then
                    oMensaje.MostrarFormulario "No puede modificar el canje de la nota de crédito de otra gestión de caja.", lcTituloNotaCredito
                    Set oMensaje = Nothing
                    Exit Function
        End If
    End If
    ValidarReglasNota = True
End Function

Public Function ValidarDatosObligNota() As Boolean
    Dim sMensaje As String
    ValidarDatosObligNota = False
    
    If Trim(txtSerieNota.Text) = "" Or Trim(txtDocumentoNota.Text) = "" Then
        sMensaje = sMensaje + "Ingrese el número de serie y documento de nota de crédito." + Chr(13)
    End If

    If sMensaje <> "" Then
       'MsgBox sMensaje, vbInformation, Me.Caption
       Dim oMensaje As New SIGHNegocios.clMensaje
       oMensaje.MostrarFormulario sMensaje, lcTituloNotaCredito
       Set oMensaje = Nothing
       Exit Function
    End If
    ValidarDatosObligNota = True
End Function

Private Sub optCanjeNotaCredito_KeyUp(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNotaDocumento_KeyUp(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNotaDocumento
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNotaSerie_KeyUp(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNotaSerie
    AdministrarKeyPreview KeyCode
End Sub

Private Sub btnAceptarPagoNota_KeyUp(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub


Sub PermisosAccesoNotaCredito()
    'PERMISOS
    Dim oRsPermisos As New Recordset
    Set oRsPermisos = mo_ReglasSeguridad.UsuariosRolesSeleccionarPermisosTodos(ml_idUsuario)
    UserControl.tabGestionCaja.TabVisible(2) = False 'Frank 24082015
    If oRsPermisos.RecordCount > 0 Then
       Do While Not oRsPermisos.EOF
          Select Case oRsPermisos.Fields!IdPermiso
          Case 207    'Caja - Ver TAB ''Devolución por Nota de Crédito Frank 24082015
               UserControl.tabGestionCaja.TabVisible(2) = True 'Frank 24082015
          End Select
          oRsPermisos.MoveNext
       Loop
    End If
    Set oRsPermisos = Nothing
End Sub

'debb-18/05/2016
Function ImprimeVariasBoletasSiPasaTopeDeItems() As Boolean
    ImprimeVariasBoletasSiPasaTopeDeItems = False
    If UCase(lcBuscaParametro.SeleccionaFilaParametro(229)) = "S" Then
       Select Case mi_Opcion
       Case sghopcionespago.sghNuevoPagoSinHistoria, sghNuevoPagoSinHistoria, sghPagarOrdenExistente
            '****** si pasa del Tope Máximo de Items, Grabar varias CABECERA/DETALLE
            '****** grabar BOLETA
    
            '****** imprimir UNA o VARIAS BOLETAS
       End Select
       ImprimeVariasBoletasSiPasaTopeDeItems = True
    End If
End Function

Sub CajaUsaDescripcionLarga()
    If lbEstaCajaUsaDescripcionLarga = True Then
       optNuevoOrdenPagoConHistoria.Enabled = False
       optOrdenExistenteFS.Enabled = False
       optPagarCtaTotal.Enabled = False
       mo_Formulario.HabilitarDeshabilitar FraServHosp, False
       optNuevoOrdenPagoSinHistoria.Value = True
       ucGestionCajaFact1.Visible = True
       ucGestionCajaFact1.Top = tabFactProductosPorCuenta.Top - 400
       ucGestionCajaFact1.Height = tabFactProductosPorCuenta.Height
       ucGestionCajaFact1.lnIGV = lnIGV
       UserControl.txtObservaciones.MaxLength = 250
    ElseIf lbTienePermisoSoloParaBoletaFarmacia = False Then
       optNuevoOrdenPagoConHistoria.Enabled = True
       optOrdenExistenteFS.Enabled = True
       optPagarCtaTotal.Enabled = True
       mo_Formulario.HabilitarDeshabilitar FraServHosp, True
       optNuevoOrdenPagoSinHistoria.Value = False
       ucGestionCajaFact1.Visible = False
       UserControl.txtObservaciones.MaxLength = 150
    End If
End Sub


Function ChequeaQueCuentaEsPagoSoloDeCita(lnIdCuenta As Long) As Boolean
  ChequeaQueCuentaEsPagoSoloDeCita = False
  On Error GoTo ChkCta
  Dim oRsTmp86 As New Recordset
  Dim oConexion As New Connection
  Dim lnIdOrden86 As Long
  oConexion.CommandTimeout = 900
  oConexion.CursorLocation = adUseClient
  oConexion.Open sighEntidades.CadenaConexion
  Set oRsTmp86 = mo_ReglasFacturacion.FactOrdenServicioSeleccionarPorIdCuenta(lnIdCuenta)
  If oRsTmp86.RecordCount = 1 Then
     If oRsTmp86!idPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaAdmisionCE And oRsTmp86!idestadofacturacion = 1 Then
        lnIdOrden86 = oRsTmp86!IdOrden
        oRsTmp86.Close
        Set oRsTmp86 = mo_ReglasFacturacion.FactOrdenServicioPagosSeleccionarPorIdOrden(lnIdOrden86, oConexion)
        If oRsTmp86.RecordCount > 0 Then
            If MsgBox("La CUENTA es de CONSULTA EXTERNA, aún no ha pagado la CITA" & Chr(13) & Chr(13) & _
                      "¿Pagará la CITA ?", vbQuestion + vbYesNo, "") <> vbNo Then
                optOrdenExistenteFS.Value = True
                cmbOrdenes.Text = oRsTmp86!IdOrdenPago
                cmbOrdenes_KeyPress 13
                ChequeaQueCuentaEsPagoSoloDeCita = True
            End If
        End If
     End If
  End If
  oRsTmp86.Close
  oConexion.Close
ChkCta:
  Set oRsTmp86 = Nothing
  Set oConexion = Nothing
End Function


Sub PagaNotaCreditoAutomaticamente(lcSerieNC As String, lcDocumentoNC As String, lnIdCaja As Long, lnIdGestionCaja As Long, _
                                   lnIdTurno As Long)
On Error GoTo errPNCA
    lbLaDevolucionNCesAutomatica = True
    lbTieneLicenciaParaNotaCreditoYsunat = True
    '
    
    AutomNC_lnIdCaja = lnIdCaja
    AutomNC_ml_IdUsuario = sighEntidades.Usuario
    AutomNC_lnIdGestionCaja = lnIdGestionCaja
    AutomNC_lnIdTurno = lnIdTurno
    '
    txtNotaSerie.Text = lcSerieNC
    txtNotaDocumento.Text = lcDocumentoNC
    BuscarNotaCredito
    btnAceptarPagoNota_Click
    Exit Sub
errPNCA:
    MsgBox Err.Description
End Sub
