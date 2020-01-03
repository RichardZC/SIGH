VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.UserControl ucGestionDevolucion 
   ClientHeight    =   10830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15960
   KeyPreview      =   -1  'True
   ScaleHeight     =   10830
   ScaleWidth      =   15960
   Begin TabDlg.SSTab tabGestionCaja 
      Height          =   7785
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   13185
      _ExtentX        =   23257
      _ExtentY        =   13732
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
      TabCaption(0)   =   "Gestión de caja"
      TabPicture(0)   =   "ucGestionDevolucion.ctx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblNombre"
      Tab(0).Control(1)=   "lblTotAnulados"
      Tab(0).Control(2)=   "lblTotFacturas"
      Tab(0).Control(3)=   "lblTotBoletas"
      Tab(0).Control(4)=   "lblNroAnulados"
      Tab(0).Control(5)=   "lblNroDocumentos"
      Tab(0).Control(6)=   "lblNroFacturas"
      Tab(0).Control(7)=   "lblNroBoletas"
      Tab(0).Control(8)=   "txtTotalCajero"
      Tab(0).Control(9)=   "Frame"
      Tab(0).Control(10)=   "fraBusqueda"
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Registro de Devoluciones"
      TabPicture(1)   =   "ucGestionDevolucion.ctx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label4"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "UcFacturacionContado1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "ucFacturacionProductos"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Frame3"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Frame5"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "fraOpciones"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Frame1"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Frame2"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "tabFactProductosPorCuenta"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Frame4"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtRazonSocial"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtNroSerie"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtNroDocumento"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).ControlCount=   16
      Begin VB.Frame fraBusqueda 
         Caption         =   "Búsqueda"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   -74880
         TabIndex        =   52
         Top             =   1080
         Width           =   12525
         Begin VB.CommandButton cmdSinApellidoMaterno 
            Caption         =   "..."
            Height          =   315
            Left            =   4440
            TabIndex        =   59
            ToolTipText     =   "Sin apellido MATERNO"
            Top             =   440
            Width           =   255
         End
         Begin VB.CommandButton cmdSinApellidoPaterno 
            Caption         =   "..."
            Height          =   315
            Left            =   2880
            TabIndex        =   58
            ToolTipText     =   "Sin apellido PATERNO"
            Top             =   440
            Width           =   255
         End
         Begin VB.TextBox txtComprobante 
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
            MaxLength       =   8
            TabIndex        =   57
            Top             =   450
            Width           =   1395
         End
         Begin VB.TextBox txtApellidoMaterno 
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
            Left            =   3240
            MaxLength       =   40
            TabIndex        =   56
            Top             =   430
            Width           =   1185
         End
         Begin VB.CommandButton btnLimpiar 
            Caption         =   "Limpiar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   11100
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   420
            Width           =   1275
         End
         Begin VB.CommandButton btnBuscar 
            Caption         =   "Buscar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9720
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   420
            Width           =   1305
         End
         Begin VB.TextBox txtApellidoPaterno 
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
            Left            =   1560
            MaxLength       =   40
            TabIndex        =   53
            Top             =   430
            Width           =   1275
         End
         Begin MSMask.MaskEdBox txtFecha1 
            Height          =   315
            Left            =   5280
            TabIndex        =   61
            Top             =   480
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
         Begin MSMask.MaskEdBox txtFecha2 
            Height          =   315
            Left            =   7200
            TabIndex        =   62
            Top             =   480
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
            Caption         =   "Comprobante      Apellido paterno      Apellido materno           Fechas Orden Pago"
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
            Index           =   0
            Left            =   120
            TabIndex        =   60
            Top             =   200
            Width           =   8715
         End
      End
      Begin VB.Frame Frame 
         Height          =   4695
         Left            =   -74880
         TabIndex        =   48
         Top             =   2040
         Width           =   12615
         Begin TabDlg.SSTab tabDevoluciones 
            Height          =   4875
            Left            =   0
            TabIndex        =   49
            Top             =   -240
            Width           =   12765
            _ExtentX        =   22516
            _ExtentY        =   8599
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   1
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
            TabCaption(0)   =   "Devoluciones"
            TabPicture(0)   =   "ucGestionDevolucion.ctx":0038
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "grdDevolucion"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Tab 1"
            TabPicture(1)   =   "ucGestionDevolucion.ctx":0054
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grdTab"
            Tab(1).ControlCount=   1
            Begin UltraGrid.SSUltraGrid grdDevolucion 
               Height          =   3810
               Left            =   120
               TabIndex        =   50
               Top             =   720
               Width           =   12315
               _ExtentX        =   21722
               _ExtentY        =   6720
               _Version        =   131072
               GridFlags       =   17040388
               UpdateMode      =   2
               LayoutFlags     =   67108884
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "grdDevolucion"
            End
            Begin UltraGrid.SSUltraGrid grdTab 
               Height          =   3930
               Left            =   -74880
               TabIndex        =   51
               Top             =   420
               Width           =   12315
               _ExtentX        =   21722
               _ExtentY        =   6932
               _Version        =   131072
               GridFlags       =   17040388
               UpdateMode      =   2
               LayoutFlags     =   67108884
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "grdTab"
            End
         End
      End
      Begin VB.TextBox txtNroDocumento 
         Height          =   285
         Left            =   5640
         TabIndex        =   47
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtNroSerie 
         Height          =   285
         Left            =   4560
         TabIndex        =   46
         Top             =   1080
         Width           =   735
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
         Left            =   120
         MaxLength       =   30
         TabIndex        =   43
         Top             =   2040
         Width           =   7020
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
         Height          =   945
         Left            =   10440
         TabIndex        =   37
         Top             =   6480
         Width           =   2655
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
            Left            =   1320
            TabIndex        =   38
            Top             =   240
            Width           =   1185
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
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   1020
         End
      End
      Begin TabDlg.SSTab tabFactProductosPorCuenta 
         Height          =   3540
         Left            =   30
         TabIndex        =   18
         Top             =   2760
         Width           =   12930
         _ExtentX        =   22807
         _ExtentY        =   6244
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
         TabPicture(0)   =   "ucGestionDevolucion.ctx":0070
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "ucFactServiciosPorCuenta"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Farmacia"
         TabPicture(1)   =   "ucGestionDevolucion.ctx":008C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label26"
         Tab(1).Control(1)=   "Label27"
         Tab(1).Control(2)=   "ucFactBienesPorCuenta"
         Tab(1).Control(3)=   "txtCtaFarmExonerado"
         Tab(1).Control(4)=   "txtCtaFarmTfarmacia"
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
            TabIndex        =   25
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
            TabIndex        =   24
            Top             =   2910
            Width           =   1200
         End
         Begin SISGalenPlus.ucFactItemsPorCuenta2 ucFactServiciosPorCuenta 
            Height          =   2520
            Left            =   105
            TabIndex        =   63
            Top             =   465
            Width           =   12570
            _ExtentX        =   22172
            _ExtentY        =   4445
         End
         Begin SISGalenPlus.ucFactItemsPorCuenta2 ucFactBienesPorCuenta 
            Height          =   2460
            Left            =   -74895
            TabIndex        =   64
            Top             =   390
            Width           =   12600
            _ExtentX        =   22225
            _ExtentY        =   4339
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
            TabIndex        =   23
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
            TabIndex        =   22
            Top             =   2970
            Width           =   1065
         End
      End
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   4275
         Begin VB.OptionButton optAnularDevoluciones 
            Caption         =   "ANULAR DEVOLUCION"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2160
            TabIndex        =   36
            Top             =   480
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.OptionButton optDevoluciones 
            Caption         =   "DEVOLUCIONES"
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
            Left            =   360
            TabIndex        =   34
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "F1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   35
            Top             =   480
            Width           =   255
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
         Height          =   2085
         Left            =   7800
         TabIndex        =   9
         Top             =   600
         Width           =   5040
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
            Left            =   1080
            MaxLength       =   7
            TabIndex        =   21
            Top             =   1440
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
            Left            =   240
            MaxLength       =   3
            TabIndex        =   20
            Top             =   1440
            Width           =   675
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
            Left            =   240
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   2
            Top             =   600
            Width           =   1845
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
            Height          =   795
            Left            =   3600
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   600
            Width           =   885
         End
         Begin VB.Label Label 
            Caption         =   "Historia Clinica"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   45
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lblOrden 
            BackStyle       =   0  'Transparent
            Caption         =   "N° Serie      N° Documento                       "
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
            Left            =   240
            TabIndex        =   16
            Top             =   1200
            Width           =   5265
         End
      End
      Begin VB.Frame fraOpciones 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   6630
         TabIndex        =   17
         Top             =   60
         Width           =   4485
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
            Picture         =   "ucGestionDevolucion.ctx":00A8
         End
         Begin Threed.SSOption optServicios 
            Height          =   345
            Left            =   690
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
            Picture         =   "ucGestionDevolucion.ctx":067C
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
         Left            =   1800
         TabIndex        =   10
         Top             =   6480
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
            TabIndex        =   12
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
            TabIndex        =   11
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
            TabIndex        =   4
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
            TabIndex        =   15
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
            TabIndex        =   14
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
            TabIndex        =   13
            Top             =   300
            Width           =   1350
         End
      End
      Begin VB.Frame Frame3 
         Height          =   915
         Left            =   120
         TabIndex        =   8
         Top             =   6480
         Width           =   1590
         Begin VB.CommandButton btnAceptar 
            Caption         =   "Aceptar (F2)"
            DisabledPicture =   "ucGestionDevolucion.ctx":0C9B
            DownPicture     =   "ucGestionDevolucion.ctx":10FB
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
            Left            =   0
            Picture         =   "ucGestionDevolucion.ctx":1570
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   120
            Width           =   1575
         End
      End
      Begin SISGalenPlus.ucFacturacionItems ucFacturacionProductos 
         Height          =   3540
         Left            =   15
         TabIndex        =   65
         Top             =   2760
         Width           =   12855
         _ExtentX        =   22781
         _ExtentY        =   7303
      End
      Begin SISGalenPlus.UcFacturacionContado UcFacturacionContado1 
         Height          =   2865
         Left            =   15
         TabIndex        =   66
         Top             =   2790
         Visible         =   0   'False
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   5054
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
         Left            =   5400
         TabIndex        =   44
         Top             =   1080
         Width           =   105
      End
      Begin VB.Label Label3 
         Caption         =   "Nº Documento"
         Height          =   255
         Left            =   5640
         TabIndex        =   42
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Serie"
         Height          =   255
         Left            =   4560
         TabIndex        =   41
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label 
         Caption         =   "Razón Social  o Apell.Paterno Apell.Materno  Nombre"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   40
         Top             =   1680
         Width           =   4095
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
         Left            =   -62370
         TabIndex        =   33
         Top             =   8460
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
         Left            =   -72780
         TabIndex        =   32
         Top             =   8850
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
         Left            =   -69600
         TabIndex        =   31
         Top             =   8850
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
         Left            =   -62370
         TabIndex        =   30
         Top             =   8850
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
         Left            =   -66240
         TabIndex        =   29
         Top             =   8850
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
         Left            =   -72780
         TabIndex        =   28
         Top             =   8460
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
         Left            =   -69600
         TabIndex        =   27
         Top             =   8460
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
         Left            =   -66240
         TabIndex        =   26
         Top             =   8460
         Width           =   180
      End
      Begin VB.Label lblNombre 
         BackColor       =   &H00373842&
         Caption         =   "Gestion de Devolución"
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
         TabIndex        =   7
         Top             =   390
         Width           =   12945
      End
   End
End
Attribute VB_Name = "ucGestionDevolucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para mantenimiento de CAJA (Emisión de Boletas, Tickets,...)
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Formulario As New sighentidades.Formulario
Dim mi_OpcionINO As sghopcionespago

Dim mo_Teclado As New sighentidades.Teclado
Dim lcBuscaParametro As New SIGHDatos.Parametros
'
Dim mo_doCajaGestion As DOCajaGestion
Dim mo_DOComprobantePago As New DOCajaComprobantesPago
Dim mo_DOComprobantePagoDevolucion As New DOCajaComprobantesPago
Dim mo_oComprobantepago As New CajaComprobantesPago
Dim mo_DOFactOrdenServicio As New DOFactOrdenServicio
Dim mo_DOFactOrdenBienInsumo As New DoFactOrdenesBienes
Dim mo_DoAtencion As New DOAtencion
Dim mo_DoFactOrdenServPagos  As New DoFactOrdenServPagos
Dim mo_DOCuentaAtencion As DOCuentaAtencion
Dim doCajero As New SIGHComun.DOCajaCajero
Dim mo_DoPaciente As New doPaciente
'
'Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
'Dim mo_AdminComun As New SIGHNegocios.ReglasComunes
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
''
'Dim mo_cmbIdPuntoCarga As New sighEntidades.ListaDespleglable
'Dim mo_cmbIdEstado As New sighEntidades.ListaDespleglable
'Dim mo_cmbFechaIngreso As New sighEntidades.ListaDespleglable
'Dim mo_cmbIdTipoGenHistoriaClinica As New sighEntidades.ListaDespleglable
'Dim mo_cmbIdCaja As New ListaDespleglable
'Dim mo_cmbIdTurno As New ListaDespleglable
'Dim mo_cmbIdCajaBusqueda As New ListaDespleglable
'Dim mo_cmbIdTurnoBusqueda As New ListaDespleglable
'Dim mo_cmbIdTipoComprobante As New ListaDespleglable
'Dim mo_cmbOrdenes As New ListaDespleglable
'Dim mo_cmbIdResponsable As New sighEntidades.ListaDespleglable
'Dim mo_cmbIdTipoFinanciamiento As New sighEntidades.ListaDespleglable
'Dim mo_cmbServicioSocial As New sighEntidades.ListaDespleglable
'
Dim oRsBusquedaRecibos As New ADODB.Recordset

'
Dim ml_IdOrdenDespacho As Long
Dim ml_IdPaciente As Long
Dim ml_IdTipoFinanciamiento As Long
Dim md_Total As Double
Dim md_Ingresado As Double
Dim md_PendientePago As Double
Dim md_PagoACuenta As Double
Dim md_Exonerado As Double
Dim ml_TipoProductoINO As Long
Dim ml_idUsuario As Long
Dim ml_PuntoCarga As Long
Dim ml_idOrden As Long
Dim ms_MensajeError As String
Dim mb_ExistenDatos As Boolean
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim ml_idCuentaAtencion As Long
Const ID_TIPO_COMPROBANTE_FACTURA = 2
Dim ml_IdGestionCaja As Long
Dim lbEsDevolucion As Boolean, lbItemEsDevolucion As Boolean
Dim ml_NombreCajero As String
Dim lnParametrosImprimeBoleta As sghImpresion
Dim ml_IdFormaPago As Long
Dim ml_IdFarmacia As Long
Dim ml_idPreVenta As Long
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
Dim lnidReceta As Long, lbSeAperturoCAJA As Boolean
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim ml_TipoProducto As Integer
Dim ml_idOrdenSeleccionado As Long
Dim mb_lbEstoyEnTabServicio As Boolean

'/************************************************/
'/*********************INO************************/
'/************************************************/
Dim mo_DOCajaDevolucion As New DOCajaDevoluciones
Dim mo_DOCajaComprobante As New DOCajaComprobantesPago
Dim lblboletadeservicios As Boolean
Dim idComprobantePagoINO As Long
Dim idCajaDevolucion As Long
'Dim mo_CajaObservacion As New CajaDevolucionesMotivo
Dim mo_AtencionesOftalmologicas As New SIGHNegocios.ReglasTriaje

'/************************************************/
'/*********************INO************************/
'/************************************************/

Property Get lbEstoyEnTabServicio() As Boolean
   lbEstoyEnTabServicio = mb_lbEstoyEnTabServicio
End Property
Property Let idOrdenSeleccionado(lValue As Long)
   ml_idOrdenSeleccionado = lValue
End Property
Property Get idOrdenSeleccionado() As Long
   idOrdenSeleccionado = ml_idOrdenSeleccionado
End Property
Property Let TipoProducto(lValue As Long)
   ml_TipoProducto = lValue
End Property
Property Get TipoProducto() As Long
   TipoProducto = ml_TipoProducto
End Property


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

Public Function Inicializar()
 
     txtFecha1.Text = Date & " 00:01"
     txtFecha2.Text = Date & " 23:59"
    
    mo_Apariencia.ConfigurarFilasBiColores grdDevolucion, sighentidades.GrillaConFilasBicolor
   ' Set mo_cmbIdCaja.MiComboBox = cmbIdCaja
'    Set mo_cmbIdTurno.MiComboBox = cmbIdTurno
'
'    Set mo_cmbIdCajaBusqueda.MiComboBox = cmbIdCajaBusqueda
'    Set mo_cmbIdTurnoBusqueda.MiComboBox = cmbIdTurnoBusqueda
'
'    Set mo_cmbIdPuntoCarga.MiComboBox = cmbIdPuntoDeCarga
'    Set mo_cmbFechaIngreso.MiComboBox = cmbFechaIngreso
'    Set mo_cmbIdTipoGenHistoriaClinica.MiComboBox = cmbIdTipoGenHistoriaClinica
'    Set mo_cmbIdTipoComprobante.MiComboBox = cmbIdTipoComprobante
'    Set mo_cmbOrdenes.MiComboBox = cmbOrdenes
'    Set mo_cmbIdResponsable.MiComboBox = cmbIdResponsable
'    Set mo_cmbIdTipoFinanciamiento.MiComboBox = cmbIdTipoFinanciamiento
'    Set mo_cmbServicioSocial.MiComboBox = cmbServicioSocial
''
'    txtFecha1.Text = Date & " 00:01"
'    txtFecha2.Text = Date & " 23:59"
'
'    txtFechaBoleta.Text = lcBuscaParametro.RetornaFechaServidorSQL()
    
'    ConfigurarTurno
'    ConfigurarCaja
'    ConfigurarPuntosDeCarga
'    ConfigurarTiposHistoriaClinica
'    ConfigurarTipoComprobante
    ConfigurarSiSeImprimeBoleta
    ConfiguraPermisos
    '
'    mo_cmbServicioSocial.BoundColumn = "IdEmpleado"
'    mo_cmbServicioSocial.ListField = "Empleado"
'    Set mo_cmbServicioSocial.RowSource = mo_ReglasFacturacion.EmpleadosSeleccionarPorFiltro("Where idLaboraArea=" & sghAreasLaboraEmpleado.sghSeguros & " and idLaboraSubArea= 9")
    '
'    mo_Formulario.HabilitarDeshabilitar cmbIdCaja, False
'    mo_Formulario.HabilitarDeshabilitar cmbIdTurno, False
'    mo_Formulario.HabilitarDeshabilitar txtFechaApertura, False
    
'    mo_Formulario.HabilitarDeshabilitar cmbIdTipoGenHistoriaClinica, False
'    mo_Formulario.HabilitarDeshabilitar txtNroHistoria, False
'    mo_Formulario.HabilitarDeshabilitar txtNombres, False
'    mo_Formulario.HabilitarDeshabilitar cmbOrdenes, False
'    mo_Formulario.HabilitarDeshabilitar cmbIdPuntoDeCarga, False
'    mo_Formulario.HabilitarDeshabilitar cmbFechaIngreso, False
    mo_Formulario.HabilitarDeshabilitar txtNroSerie, False
    mo_Formulario.HabilitarDeshabilitar txtNroDocumento, False
  '  mo_Formulario.HabilitarDeshabilitar txtObservaciones, False
   ' mo_Formulario.HabilitarDeshabilitar txtDni, False
    
    UserControl.tabGestionCaja.TabVisible(1) = False
    
    ucFacturacionProductos.TipoProducto = sghServicio
    ml_TipoProductoINO = sghServicio
    
    ucFacturacionProductos.idUsuario = ml_idUsuario
    ucFacturacionProductos.Inicializar
    
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
    
    'Configuracion para PreVenta (FARMACIA)
    ml_idConfiguracionParaPreventa = Val(lcBuscaParametro.SeleccionaFilaParametro(229))
    '
    InicilizarParametros
    '
    On Error Resume Next
   UserControl.txtNroSerie.SetFocus
End Function


Sub InicilizarParametros()
        wxParametro102 = lcBuscaParametro.SeleccionaFilaParametro(102)
        wxParametro205 = lcBuscaParametro.SeleccionaFilaParametro(205)
        wxParametro206 = lcBuscaParametro.SeleccionaFilaParametro(206)
        wxParametro207 = lcBuscaParametro.SeleccionaFilaParametro(207)
        wxParametro211 = lcBuscaParametro.SeleccionaFilaParametro(211)
        wxParametro221 = lcBuscaParametro.SeleccionaFilaParametro(221)
        wxParametro208 = lcBuscaParametro.SeleccionaFilaParametro(208)
        wxParametro285 = lcBuscaParametro.SeleccionaFilaParametro(285)
        wxParametro286 = lcBuscaParametro.SeleccionaFilaParametro(286)
        wxParametro288 = lcBuscaParametro.SeleccionaFilaParametro(288)
        wxParametro339 = lcBuscaParametro.SeleccionaFilaParametro(339)
        wxParametro346 = lcBuscaParametro.SeleccionaFilaParametro(346)
End Sub


'***************daniel barrantes**************
'***************Retorna si se Imprime BOLETA o solo PANTALLA
'***************
Sub ConfigurarSiSeImprimeBoleta()
    lnParametrosImprimeBoleta = mo_reglasComunes.ParametrosSeleccionarValorIntPorTipoYCodigo("INDICADOR", "IMPRIMIR_RECIBO")
End Sub

Public Function RealizarAperturaDeCaja(lIdUsuario As Long, lIdCaja As Long, lIdTurno As Long, lbEmiteSoloServicios As Boolean) As Boolean
Dim oDOCajaGestion As DOCajaGestion
Dim bAperturaOK As Boolean

    bAperturaOK = False
    Set oDOCajaGestion = mo_AdminCaja.RetornaCajaAbierta(lIdUsuario, lIdCaja, lIdTurno)
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
'        mo_cmbIdTurno.BoundText = oDOCajaGestion.IdTurno
'        mo_cmbIdCaja.BoundText = oDOCajaGestion.IdCaja
     '   txtFechaApertura = oDOCajaGestion.FechaApertura
'        mo_cmbIdTipoComprobante.BoundText = wxIdTipoComprobanteDefault        'Boleta
'        mo_cmbIdPuntoCarga.BoundText = 99   'Cajero
        UserControl.tabGestionCaja.TabVisible(1) = True
        UserControl.tabGestionCaja.Tab = 1
        UserControl.KeyPreview = True   'debb-16/02/2011
    End If
    
    RealizarAperturaDeCaja = bAperturaOK
    On Error Resume Next
    UserControl.txtNroSerie.SetFocus
End Function

Public Function RealizarCierreDeCaja() As Boolean
    
    RealizarCierreDeCaja = False
    mo_doCajaGestion.fechaCierre = lcBuscaParametro.RetornaFechaHoraServidorSQL    'Now
    mo_doCajaGestion.EstadoLote = "C"
    
    If mo_AdminCaja.CajaGestionModificar(mo_doCajaGestion) Then
        UserControl.tabGestionCaja.TabVisible(1) = False
        RealizarCierreDeCaja = True
        UserControl.KeyPreview = False   'debb-16/02/2011
    End If
    
End Function




Private Sub btnAceptar_KeyUp(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub


Sub LimpiarOpciones()
    Set mo_DOFactOrdenServicio = Nothing
    Set mo_DoAtencion = Nothing
    Set mo_DOComprobantePago = Nothing
    Set ucFacturacionProductos.Atencion = Nothing
    mo_DOFactOrdenServicio.IdOrden = 0
    mo_DOFactOrdenBienInsumo.IdOrden = 0
'    mo_cmbIdTipoGenHistoriaClinica.BoundText = wxParametro211
    txtNroHistoria = ""
   ' txtNombres = ""
   ' cmbOrdenes.Text = ""
'    mo_cmbFechaIngreso.BoundText = 0
   'txtIngresado = ""
   'txtPendientePago = ""
   'txtPagoACuenta = "0"
    'txtExonerado = "0"
    md_Total = 0
    txtTotal.Text = "0"
    txtEfectivo = ""
    txtFalta = ""
    txtVuelto = ""
    txtRazonSocial = ""
   ' txtRuc = ""
    txtNserieB.Text = ""
    txtNdocumentoB.Text = ""
    ml_IdPaciente = 0
    ml_IdFormaPago = 1          'Contado
    ml_IdFarmacia = 0           '1=Farmacia Principal,2=Farmacia Emergencia,0-otros
    ml_idPreVenta = 0
    ml_idCuentaAtencion = 0
    'txtNreceta.Text = "": lnidReceta = 0
    '********INO***********
'    cmdConsulta.Visible = IIf(mi_OpcionINO = sghDevolucionINO, True, False)
    'debb-16/02/2011
   ' txtNroCuenta.Text = ""
    ml_IdOrdenDespacho = 0
    'mo_Formulario.HabilitarDeshabilitar txtExonerado, False
    UcFacturacionContado1.Visible = False
    'frmPreventaServ.Visible = False
    'chkGeneraPreventaServ.Value = 0
'    mo_cmbIdTipoFinanciamiento.BoundText = "1"
    
'    mo_Formulario.HabilitarDeshabilitar cmbIdTipoGenHistoriaClinica, False
    mo_Formulario.HabilitarDeshabilitar txtNroHistoria, False
   ' mo_Formulario.HabilitarDeshabilitar txtNombres, False
   ' mo_Formulario.HabilitarDeshabilitar cmbOrdenes, False
   ' mo_Formulario.HabilitarDeshabilitar cmbIdPuntoDeCarga, False
   'mo_Formulario.HabilitarDeshabilitar cmbFechaIngreso, False
   ' mo_Formulario.HabilitarDeshabilitar txtNroCuenta, False
    mo_Formulario.HabilitarDeshabilitar txtNserieB, False
    mo_Formulario.HabilitarDeshabilitar txtNdocumentoB, False
   ' mo_Formulario.HabilitarDeshabilitar txtPagoACuenta, False
'    mo_Formulario.HabilitarDeshabilitar cmbIdTipoFinanciamiento, False
   ' mo_Formulario.HabilitarDeshabilitar txtObservaciones, False
   ' mo_Formulario.HabilitarDeshabilitar txtDni, False    'debb-18/02/2011
    
    'cmdPaquetes.Enabled = False: lnIdFactPaquete = 0
    lbItemEsDevolucion = False: 'txtObservaciones.Text = ""
    'debb-16/02/2011
    lbCargaEstadoDeCuentaFS = False
    UserControl.KeyPreview = True
    'txtDni.Text = ""
   ' UserControl.txtCtaFarmExonerado.Text = ""
'    UserControl.txtCtaFarmTfarmacia.Text = ""
'    UserControl.txtCtaServExonerado.Text = ""
'    UserControl.txtCtaServTservicio.Text = ""
'    lblCuentaConSeguro.Caption = ""
    'debb-16/02/2011
   ' cmbServicioSocial.Visible = False: txtServicioSocial.Visible = False: lblServicioSocial.Visible = False
    'txtNreceta.Enabled = True: cmbBuscaReceta.Enabled = True
End Sub






Private Sub cmbIdTipoComprobante_Click()
Dim lIdTipoComprobante As Long
Dim oCajaNroDocumento As New DOCajaNroDocumento
Dim rsBuscaBoleta As Recordset, lbSigue As Boolean, lnLen As Integer
'
'    txtRuc.Enabled = True
'    txtNroSerie.Text = ""
    txtNroDocumento.Text = ""
'    lIdTipoComprobante = Val(mo_cmbIdTipoComprobante.BoundText)
    
    
    If lIdTipoComprobante > 0 Then
        Set oCajaNroDocumento = mo_AdminCaja.NroDocumentoSeleccionarPorIdCajaYTipoComprobante(mo_doCajaGestion.IdCaja, lIdTipoComprobante)
        txtNroSerie.Text = Trim(oCajaNroDocumento.nroSerie)
        txtNroDocumento.Text = Trim(oCajaNroDocumento.nrodocumento)
        'comprueba que no existe esa nueva Boleta
        lbSigue = True
        Do While lbSigue = True
           Set rsBuscaBoleta = mo_AdminCaja.CajaComprobantePagoPorSerieDocumento(txtNroSerie.Text, txtNroDocumento.Text)
           If rsBuscaBoleta.RecordCount = 0 Then
              lbSigue = False
           Else
              lnLen = Len(txtNroDocumento.Text)
              txtNroDocumento.Text = Right("00000000" & Trim(Str(Val(txtNroDocumento.Text) + 1)), lnLen)
           End If
        Loop
        '
'        If lIdTipoComprobante <> ID_TIPO_COMPROBANTE_FACTURA Then
'            mo_Formulario.HabilitarDeshabilitar txtRuc, False
'            lbEsUnaFactura = False
'        Else
'            mo_Formulario.HabilitarDeshabilitar txtRuc, True
'            lbEsUnaFactura = True
'        End If
    End If
    Set oCajaNroDocumento = Nothing
    '
    wxIdTipoComprobanteDefault = lIdTipoComprobante
    CargaSetup_Caja App.Path & "\archivos", wxIdTipoComprobanteDefault, False
End Sub

Private Sub cmbIdTipoComprobante_KeyUp(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub


'Private Sub cmdConsulta_Click()
' Dim oBusqueda As New CajaDevolucionesConsulta
'    oBusqueda.Show 1
'    Set oBusqueda = Nothing
'End Sub

Private Sub cmdLeer_Click()
'    ml_TipoProductoINO = sghServicio
 
'        If mi_OpcionINO = sghPagarOrdenExistenteF Then 'solo es usado cuando es "pagar orden existente"
'            LeerBienesPorTipoDePago
'        Else
'            Select Case ml_TipoProductoINO
'            Case sghServicio
                LeerServiciosPorTipoDePago
'            Case sghbien
'                LeerBienesPorTipoDePago
'            End Select
'        End If


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


Sub InabilitaTotales()
   ' mo_Formulario.HabilitarDeshabilitar txtIngresado, False
   ' mo_Formulario.HabilitarDeshabilitar txtPendientePago, False
   ' mo_Formulario.HabilitarDeshabilitar txtExonerado, False
   ' mo_Formulario.HabilitarDeshabilitar txtPagoACuenta, False
'    mo_Formulario.HabilitarDeshabilitar txtTotal, False
'    mo_Formulario.HabilitarDeshabilitar txtVuelto, False
'    mo_Formulario.HabilitarDeshabilitar txtFalta, False
'    mo_Formulario.HabilitarDeshabilitar txtCtaFarmTfarmacia, False
'    mo_Formulario.HabilitarDeshabilitar txtCtaFarmExonerado, False
'    mo_Formulario.HabilitarDeshabilitar txtCtaServTservicio, False
'    mo_Formulario.HabilitarDeshabilitar txtCtaServExonerado, False

End Sub




''debb-16/02/2011
'Private Sub tabGestionCaja_Click(PreviousTab As Integer)
'    If PreviousTab = 0 Then
'       UserControl.KeyPreview = True
'    Else
'       UserControl.KeyPreview = False
'    End If
'End Sub


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

'Private Sub txtEfectivo_KeyPress(KeyAscii As Integer)
'
'    If KeyAscii = 13 Then
'       btnAceptar.SetFocus
'       Exit Sub
'    End If
'    If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
'       If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
'           KeyAscii = 0
'       End If
'    End If
'
'End Sub


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


Private Sub txtNserieB_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNserieB
End Sub

Private Sub txtNserieB_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtNserieB_KeyUp(KeyCode As Integer, Shift As Integer)
     AdministrarKeyPreview KeyCode
End Sub

'Private Sub txtPagoACuenta_Change()
'   ActualizaTotalApagar
'   md_Total = txtTotal.Text
'End Sub


'Private Sub txtPagoACuenta_KeyDown(KeyCode As Integer, Shift As Integer)
'         AdministrarKeyPreview KeyCode
'End Sub
'
'
'Private Sub txtPagoACuenta_KeyUp(KeyCode As Integer, Shift As Integer)
'    AdministrarKeyPreview KeyCode
'End Sub
'
'Private Sub txtVuelto_KeyUp(KeyCode As Integer, Shift As Integer)
'    AdministrarKeyPreview KeyCode
'End Sub

'Private Sub ucFactBienesPorCuenta_SePresionoTeclaEspecial(KeyCode As Integer)
'     If KeyCode >= vbKeyF2 And KeyCode <= vbKeyF12 Then
'        AdministrarKeyPreview KeyCode
'     End If
'
'End Sub


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
    Case vbKeyF2
          btnAceptar_Click
    End Select
End Sub



'Public Sub ConfigurarCaja()
'
''    mo_cmbIdCaja.BoundColumn = "IdCaja"
''    mo_cmbIdCaja.ListField = "Descripcion"
''    Set mo_cmbIdCaja.RowSource = mo_AdminCaja.CajaSeleccionarTodosParaLista()
'
'    mo_cmbIdCajaBusqueda.BoundColumn = "IdCaja"
'    mo_cmbIdCajaBusqueda.ListField = "Descripcion"
'    Set mo_cmbIdCajaBusqueda.RowSource = mo_AdminCaja.CajaSeleccionarTodosParaLista()
'
'End Sub

'Public Sub ConfigurarTurno()
'
'    mo_cmbIdTurno.BoundColumn = "IdTurno"
'    mo_cmbIdTurno.ListField = "Descripcion"
'    Set mo_cmbIdTurno.RowSource = mo_AdminCaja.TurnosSeleccionarTodosParaLista()
'
'    mo_cmbIdTurnoBusqueda.BoundColumn = "IdTurno"
'    mo_cmbIdTurnoBusqueda.ListField = "Descripcion"
'    Set mo_cmbIdTurnoBusqueda.RowSource = mo_AdminCaja.TurnosSeleccionarTodosParaLista()
'
'    mo_cmbIdResponsable.BoundColumn = "IdEmpleado"
'    mo_cmbIdResponsable.ListField = "DCajero"
'    Set mo_cmbIdResponsable.RowSource = mo_AdminCaja.CajerosSeleccionarTodos()
'
'End Sub


'********************************************************************************************
'********************************************************************************************
'********************************************************************************************
'                                   COMPROBANTE DE PAGO
'********************************************************************************************
'********************************************************************************************
'********************************************************************************************

Private Sub btnAceptar_Click()

    If CCur(txtTotal.Text) < 0 Then
       MsgBox "El total es menor a CERO", vbInformation, "CAJA"
       Exit Sub
    End If
    If mi_OpcionINO = sghNuevoPagoConHistoria Or mi_OpcionINO = sghNuevoPagoSinHistoria Then
       lnTotalGrid = ucFacturacionProductos.DevuelveTotalPagar
'       ActualizaTotalApagar
    End If
    If MsgBox("Por favor confirmar, ¿Realmente desea grabar los cambios que ha realizado?", vbCritical + vbYesNo, "Estado de Cuenta") = vbNo Then
        Exit Sub
    End If
   
    Select Case mi_OpcionINO

    '*************************************************
    '************************INO**********************
    '*************************************************
    Case sghopcionespago.sghDevolucionINO
    If MsgBox("Por favor confirmar, La cita ligada al comprobante de pago será anulada. ¿Está seguro de continuar con la operación?", vbCritical + vbYesNo, "Estado de Cuenta") = vbNo Then
        Exit Sub
    Else
        If ValidarReglas() Then
'               mo_CajaObservacion.Show 1
               'MsgBox "Porfavor ingrese el motivo por el cual se origina la devolucion"
               If DevolucionINO() Then
                    MsgBox "La orden se ha devuelto y el cupo se ha liberado correctamente", vbInformation, "Gestión de Caja"
                    MsgBox "N° Devolución: " & mo_DOCajaDevolucion.idDevolucion, vbInformation, ""
                    cmbIdTipoComprobante_Click
                Else
                    MsgBox "No se pudo generar la devolucion"
               End If
           End If
    End If
     
     Case sghopcionespago.sghAnularDevolucionINO
     If ValidarReglas() Then
               If AnularDevolucionINO() Then
                    MsgBox "La devolucion de la orden ha sido anulada correctamente", vbInformation, "Gestión de Caja"
                    MsgBox "Se ha activado el comprobante de pago"
                    cmbIdTipoComprobante_Click
                Else
                    MsgBox "No se pudo generar la devolucion"
               End If
           End If
    '*************************************************
    '************************INO**********************
    '*************************************************
    End Select
    
    LimpiarFormulario
    LimpiarOpciones
    mo_Formulario.HabilitarDeshabilitar txtNserieB, True
    mo_Formulario.HabilitarDeshabilitar txtNdocumentoB, True
End Sub


Sub ImprimirComprobanteServicio()
Dim oRptCaja As New RptCaja
End Sub

Sub ImprimirComprobanteBienesInsumos()
Dim oRptCaja As New RptCaja

    If mo_reglasComunes.ParametrosSeleccionarValorIntPorTipoYCodigo("INDICADOR", "IMPRIMIR_RECIBO") = 1 Then
        oRptCaja.ImprimirComprobantePagoBienesInsumos mo_DoPaciente, mo_DoAtencion, mo_DOFactOrdenBienInsumo, mo_DOComprobantePago, ucFacturacionProductos.FacturacionProductos
    End If
    
End Sub


Sub LimpiarFormulario()
    Select Case mi_OpcionINO
'    Case sghPagarCuentaExistente, sghPagarCuentaTotalFS    'debb-17/02/2011
'        ucFactServiciosPorCuenta.LimpiarGrilla
'        ucFactBienesPorCuenta.LimpiarGrilla
        
    Case sghopcionespago.sghDevolucionINO, sghopcionespago.sghAnularDevolucionINO
        ucFacturacionProductos.FiltraCpt = sghMuestraTodosCpt
        ucFacturacionProductos.LimpiarGrilla
        ucFacturacionProductos.PermiteAgregarItems = False
        UcFacturacionContado1.LimpiarGrilla
        ucFacturacionProductos.LimpiarGrilla
        ucFacturacionProductos.Visible = True
        
        ucFacturacionProductos.Visible = True
    End Select
End Sub

Function ValidarReglas() As Boolean
   ValidarReglas = False
      
   ValidarReglas = True
End Function
'Function AgregarDatos() As Boolean
'    Dim oDllFactUCGestionCaja As New SighFacturacion.dllFactUCGestionCaja
'    If mi_OpcionINO = sghPagarOrdenExistenteF Then 'solo es usado cuando es "pagar orden existente"
'        '
'        'AgregarDatos = mo_ReglasFarmacia.CajaComprobantePagoBienesRegistraBoleta(mo_DOComprobantePago, mo_doCajaGestion, mo_DOFactOrdenBienInsumo, UcFacturacionContado1.DevuelveProductos, ml_idUsuario, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtRazonSocial.Text)
'        AgregarDatos = oDllFactUCGestionCaja.CajaComprobantePagoBienesRegistraBoleta(mo_DOComprobantePago, mo_doCajaGestion, mo_DOFactOrdenBienInsumo, UcFacturacionContado1.DevuelveProductos, ml_idUsuario, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtRazonSocial.Text)
'        '
'        ms_MensajeError = mo_ReglasFarmacia.MensajeError
'    Else
'        Select Case ml_TipoProductoINO
'        Case sghServicio
'            If lnIdFactPaquete > 0 Then
'               AgregarDatos = mo_AdminCaja.CajaComprobantePagoServicioPaqueteAgregar(mo_DOComprobantePago, mo_doCajaGestion, mo_DoFactOrdenServPagos, ucFacturacionProductos.FacturacionProductos, ml_idUsuario, mo_DoAtencion, Val(mo_cmbIdPuntoCarga.BoundText), mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
'            Else
''               AgregarDatos = mo_AdminCaja.CajaComprobantePagoServicioAgregar(mo_DOComprobantePago, mo_doCajaGestion, _
''                              mo_DoFactOrdenServPagos, ucFacturacionProductos.FacturacionProductos, ml_idUsuario, _
''                              mo_DOAtencion, Val(mo_cmbIdPuntoCarga.BoundText), mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, _
''                              lnIdReceta)
'               AgregarDatos = oDllFactUCGestionCaja.CajaComprobantePagoServicioAgregar(mo_DOComprobantePago, mo_doCajaGestion, _
'                              mo_DoFactOrdenServPagos, ucFacturacionProductos.FacturacionProductos, ml_idUsuario, _
'                              mo_DoAtencion, Val(mo_cmbIdPuntoCarga.BoundText), mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, _
'                              lnidReceta)
'            End If
'        Case sghbien
'            AgregarDatos = mo_AdminCaja.CajaComprobantePagoBienInsumoAgregar(mo_DOComprobantePago, mo_doCajaGestion, mo_DOFactOrdenBienInsumo, ucFacturacionProductos.FacturacionProductos, ucFacturacionProductos.ProductosEliminados, ml_idUsuario, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
'
'        End Select
'        ms_MensajeError = mo_AdminCaja.MensajeError
'    End If
'    Set oDllFactUCGestionCaja = Nothing
'End Function



Private Sub btnCancelar_Click()
    'Visible = False
End Sub

Private Sub Form_Load()

   mb_lbEstoyEnTabServicio = False

    
    'Set mo_cmbFechaIngreso.MiComboBox = cmbFechaIngreso
    'Set mo_cmbIdPuntoCarga.MiComboBox = cmbIdPuntoDeCarga
'    Set mo_cmbIdTipoGenHistoriaClinica.MiComboBox = cmbIdTipoGenHistoriaClinica
    
'    ConfigurarPuntosDeCarga
'    ConfigurarTiposHistoriaClinica
'    ConfigurarFechaIngreso
    
'    mo_cmbIdPuntoCarga.BoundText = ml_PuntoCarga
'    mo_cmbIdTipoGenHistoriaClinica.BoundText = 2
    
    
    ucFacturacionProductos.idUsuario = ml_idUsuario
    ucFacturacionProductos.Inicializar
    ucFacturacionProductos.idTipoFinanciamiento = ml_IdTipoFinanciamiento
    ucFacturacionProductos.TipoProducto = sghServicio
    ucFacturacionProductos.idPuntoCarga = ml_PuntoCarga
    
    CargarDatosAlFormulario
    
End Sub

Sub CargarDatosAlFormulario()
End Sub

'Sub CargarDatosServiciosALosControlesPorIdOrden()
'Dim oRsTmp1 As New Recordset
'
''        If Trim(cmbOrdenes.Text) = "" Then
''            MsgBox "Ingrese el numero de la orden que desea consultar", vbInformation, "Caja"
''            Exit Sub
''        End If
'
'        ucFacturacionProductos.LimpiarGrilla
'
'        'Carga datos de la orden
'        Set oRsTmp1 = mo_AdminCaja.FactOrdenServicioPagosPendientesDePagoPorIdOrdenPago(Val(cmbOrdenes.Text))
'        If oRsTmp1.RecordCount = 0 Then
'            mb_ExistenDatos = False
'            Exit Sub
'        End If
'        mo_cmbIdPuntoCarga.BoundText = oRsTmp1.Fields!idPuntoCarga
'        IdOrden = Val(cmbOrdenes.Text)
'        ml_IdOrdenDespacho = oRsTmp1.Fields!IdOrden
'        mb_ExistenDatos = True
'        mo_DoAtencion.idAtencion = IIf(IsNull(oRsTmp1.Fields!idAtencion), 0, oRsTmp1.Fields!idAtencion)
'        mo_DoAtencion.idPaciente = IIf(IsNull(oRsTmp1.Fields!idPaciente), 0, oRsTmp1.Fields!idPaciente)
'        mo_DoAtencion.IdFormaPago = oRsTmp1.Fields!idTipoFinanciamiento
'        Set mo_DoFactOrdenServPagos = mo_ReglasFacturacion.FactOrdenServicioPagosSeleccionarPorIdOrdenPago(oRsTmp1.Fields!IdOrdenPago)
'
'           'Valida el estado de la orden
'        Select Case mi_OpcionINO
'        Case sghopcionespago.sghPagarOrdenExistente
'            Select Case oRsTmp1.Fields!idestadofacturacion
'            Case 1
'            Case 4
'                MsgBox "La orden ya ha sido PAGADA.", vbInformation, "Caja"
'                Exit Sub
'            Case 9
'                MsgBox "La orden no puede ser pagada, se encuentra en estado ANULADO.", vbInformation, "Caja"
'                Exit Sub
'            End Select
'        Case sghopcionespago.sghPagarCuentaExistente
'
'        Case sghopcionespago.sghDevolucion
'            Select Case oRsTmp1.Fields!idestadofacturacion
'            Case 1
'                MsgBox "La orden aun no ha sido PAGADA, solo se puede realizar devoluciones de ordenes PAGADAS.", vbInformation, "Caja"
'                Exit Sub
'            Case 4
'                'Solo se pueden devolver ordenes pagadas
'                'Verificar que tenga productos con autorizacion de devolucion
'            Case 9
'                MsgBox "No se puede realizar la devolución, la orden se encuentra en estado ANULADO.", vbInformation, "Caja"
'                Exit Sub
'            End Select
'        Case sghopcionespago.sghAnulacion
'            Select Case oRsTmp1.Fields!idestadofacturacion
'            Case 1
'                MsgBox "La orden aun no ha sido PAGADA, solo se puede anular de ordenes PAGADAS.", vbInformation, "Caja"
'                Exit Sub
'            Case 4
'                'Solo se pueden devolver ordenes pagadas
'                'Verificar que tenga productos con autorizacion de devolucion
'            Case 9
'                MsgBox "No se puede realizar la anulación, la orden ya se encuentra ANULADA.", vbInformation, "Caja"
'                Exit Sub
'            End Select
'        End Select
'
'        'Cargar datos del paciente y de la atencion
'        ml_idPaciente = IIf(IsNull(oRsTmp1.Fields!idPaciente), 0, oRsTmp1.Fields!idPaciente)
'        txtNombres.Text = IIf(IsNull(oRsTmp1.Fields!ApellidoPaterno), "", oRsTmp1.Fields!ApellidoPaterno + " " + oRsTmp1.Fields!ApellidoMaterno + " " + oRsTmp1.Fields!PrimerNombre)
'        txtNroHistoria.Text = IIf(IsNull(oRsTmp1.Fields!NroHistoriaClinica), 0, oRsTmp1.Fields!NroHistoriaClinica)
'        txtRazonSocial = txtNombres
'        txtNroCuenta.Text = IIf(IsNull(oRsTmp1.Fields!idCuentaAtencion), 0, oRsTmp1.Fields!idCuentaAtencion)
'        ml_idCuentaAtencion = IIf(IsNull(oRsTmp1.Fields!idCuentaAtencion), 0, oRsTmp1.Fields!idCuentaAtencion)
'        mo_cmbIdTipoFinanciamiento.BoundText = oRsTmp1.Fields!idTipoFinanciamiento
'
'        'Cargar datos de los servicios
'        Select Case mi_OpcionINO
'        Case sghopcionespago.sghPagarOrdenExistente
'        Case sghopcionespago.sghPagarCuentaExistente
'        Case sghopcionespago.sghDevolucion
'            ucFacturacionProductos.EstadosFacturacion = "5"    'Autorizados a devolver
'            ucFacturacionProductos.TiposFinanciamiento = "1,5,9"
'        Case sghopcionespago.sghAnulacion
'            ucFacturacionProductos.EstadosFacturacion = "4"    'Pagados
'            ucFacturacionProductos.TiposFinanciamiento = "1,5,9"
'        End Select
'        ucFacturacionProductos.IdOrdenPago = oRsTmp1.Fields!IdOrdenPago
'        ucFacturacionProductos.CargaProductosPorIdOrdenPago
'        txtExonerado.Text = mo_DoFactOrdenServPagos.ImporteExonerado
'        If mi_OpcionINO <> sghPagarOrdenExistente Then    'debb-13/03/2012
'           lnTotalGrid = md_Total + Val(txtExonerado.Text)
'        End If
'        ActualizaTotalApagar
'        oRsTmp1.Close
'        Set oRsTmp1 = Nothing
'End Sub

'Sub CargarDatosBienesALosControlesPorIdOrden()
'Dim oDOFactOrdenBienInsumo As DoFactOrdenesBienes
'Dim oDoPreVenta As New DoFarmPreVenta
'Dim oPreVenta As New FarmPreVenta
'Dim oConexion As New ADODB.Connection
'        oConexion.Open sighEntidades.CadenaConexion
'        oConexion.CursorLocation = adUseClient
'        cmbOrdenes.Text = IIf(InStr(cmbOrdenes.Text, " ") > 0, Mid(cmbOrdenes.Text, 1, InStr(cmbOrdenes.Text, " ")), cmbOrdenes.Text)
'        If Trim(cmbOrdenes.Text) = "" Then
'            MsgBox "Ingrese el numero de la orden que desea consultar", vbInformation, "Caja"
'            Exit Sub
'        End If
'        IdOrden = Val(cmbOrdenes.Text)
'        txtTotal.Text = ""
'
'        'Carga datos de la orden
'        Select Case mi_OpcionINO
'        Case sghopcionespago.sghPagarOrdenExistenteF
'             UcFacturacionContado1.InHabilitaEdicionColumnasDelGrid = True
'             UcFacturacionContado1.inicializar
'             UcFacturacionContado1.idPreVenta = 0
'             UcFacturacionContado1.LimpiarGrilla
'             Set oDOFactOrdenBienInsumo = mo_ReglasFacturacion.FactOrdenesBienesInsumoSeleccionarPorIdPreVenta(Val(cmbOrdenes.Text))
'        Case Else
'             ucFacturacionProductos.LimpiarGrilla
'             Set oDOFactOrdenBienInsumo = mo_ReglasFacturacion.FactOrdenesBienesInsumoSeleccionarPorId(Val(cmbOrdenes.Text))
'        End Select
'        If Not oDOFactOrdenBienInsumo Is Nothing Then
'             ml_idCuentaAtencion = oDOFactOrdenBienInsumo.idCuentaAtencion
'             ml_idPaciente = oDOFactOrdenBienInsumo.idPaciente
'             ml_IdFormaPago = 1  'contado
'             ml_IdTipoFinanciamiento = ml_IdFormaPago
'             oDoPreVenta.idPreVenta = oDOFactOrdenBienInsumo.idPreVenta
'             ml_idPreVenta = oDOFactOrdenBienInsumo.idPreVenta
'             Set oPreVenta.Conexion = oConexion
'             If Not oPreVenta.SeleccionarPorId(oDoPreVenta) Then
'                MsgBox "Problemas con tabla PRE-VENTA", vbInformation, "Caja"
'                Exit Sub
'             End If
'             ml_IdFarmacia = oDoPreVenta.IdAlmacen
'             Set mo_DOFactOrdenBienInsumo = oDOFactOrdenBienInsumo
'             With mo_DOFactOrdenBienInsumo
'                mo_cmbIdPuntoCarga.BoundText = mo_DOFactOrdenBienInsumo.idPuntoCarga
'
'                 mb_ExistenDatos = True
'             End With
'         Else
'            mb_ExistenDatos = False
'            Exit Sub
'         End If
'
'           'Valida el estado de la orden
'        Select Case mi_OpcionINO
'        Case sghopcionespago.sghPagarOrdenExistenteF
'            Select Case mo_DOFactOrdenBienInsumo.idestadofacturacion
'            Case 1
'            Case 4
'                MsgBox "La orden ya ha sido PAGADA.", vbInformation, "Caja"
'                Exit Sub
'            Case 9
'                MsgBox "La orden no puede ser pagada, se encuentra en estado ANULADO.", vbInformation, "Caja"
'                Exit Sub
'            End Select
'            If ml_IdFormaPago <> 1 Then
'                MsgBox "Solo se puede EMITIR BOLETA para DOCUMENTO=CONTADO", vbInformation, "Caja"
'                Exit Sub
'            End If
'        Case sghopcionespago.sghPagarCuentaExistente
'
'        Case sghopcionespago.sghDevolucion
'            Select Case mo_DOFactOrdenBienInsumo.idestadofacturacion
'            Case 1
'                MsgBox "La orden aun no ha sido PAGADA, solo se puede realizar devoluciones de ordenes PAGADAS.", vbInformation, "Caja"
'                Exit Sub
'            Case 4
'                'Solo se pueden devolver ordenes pagadas
'                'Verificar que tenga productos con autorizacion de devolucion
'            Case 9
'                MsgBox "No se puede realizar la devolución, la orden se encuentra en estado ANULADO.", vbInformation, "Caja"
'                Exit Sub
'            End Select
'
'        Case sghopcionespago.sghAnulacion
'            Select Case mo_DOFactOrdenBienInsumo.idestadofacturacion
'            Case 1
'                MsgBox "La orden aun no ha sido PAGADA, solo se puede realizar anulaciones de ordenes PAGADAS.", vbInformation, "Caja"
'                Exit Sub
'            Case 4
'                'Solo se pueden devolver ordenes pagadas
'                'Verificar que tenga productos con autorizacion de devolucion
'            Case 9
'                MsgBox "No se puede realizar la anulación, la orden se encuentra en estado ANULADO.", vbInformation, "Caja"
'                Exit Sub
'            End Select
'
'        End Select
'
'
'        'Cargar datos del paciente y de la atencion
'        If ml_idPaciente > 0 Then
''            Set mo_DOAtencion = mo_AdminAdmision.AtencionesSeleccionarPorId(mo_DOFactOrdenBienInsumo.idAtencion)
''            Set ucFacturacionProductos.Atencion = mo_DOAtencion
''            cmbFechaIngreso.Text = mo_DOAtencion.FechaIngreso
'        End If
'        txtNombres.Text = ""
'        mo_cmbIdTipoGenHistoriaClinica.BoundText = ""
'        txtNroHistoria.Text = ""
'        txtRazonSocial = ""
'        If ml_idPaciente > 0 Then
'            Dim oDOCajaComprobantesPago As New DOCajaComprobantesPago
'            Set mo_DOComprobantePago = mo_AdminCaja.ComprobantePagoSeleccionarPorId(mo_DOFactOrdenBienInsumo.IdComprobantePago, oConexion)
'            Dim oDOPaciente As New doPaciente
'            Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(ml_idPaciente, oConexion)
'            If Not oDOPaciente Is Nothing Then
'                txtNombres.Text = oDOPaciente.ApellidoPaterno + " " + oDOPaciente.ApellidoMaterno + " " + oDOPaciente.PrimerNombre
'                mo_cmbIdTipoGenHistoriaClinica.BoundText = oDOPaciente.IdTipoNumeracion
'                txtNroHistoria.Text = oDOPaciente.NroHistoriaClinica
'                txtRazonSocial = txtNombres
'            End If
'        Else
'        End If
'        'Cargar datos de los BienInsumos
'        Select Case mi_OpcionINO
'        Case sghopcionespago.sghPagarOrdenExistenteF
'
'            UcFacturacionContado1.Visible = True
'            ucFacturacionProductos.Visible = False
'
'            UcFacturacionContado1.IdAlmacen = ml_IdFarmacia
'            UcFacturacionContado1.idPreVenta = oDOFactOrdenBienInsumo.idPreVenta
'            UcFacturacionContado1.CargaProductosPorIdPreVenta
'            lnTotalGrid = oDoPreVenta.Total
'            txtExonerado.Text = oDOFactOrdenBienInsumo.ImporteExonerado
'            ActualizaTotalApagar
'            md_Total = CDbl(txtTotal.Text)
'        Case sghopcionespago.sghPagarCuentaExistente
'            UcFacturacionContado1.Visible = False
'            ucFacturacionProductos.Visible = True
'            ucFacturacionProductos.IdOrden = IdOrden
'            ucFacturacionProductos.CargaProductosPorIdOrden
'        Case sghopcionespago.sghDevolucion
'            UcFacturacionContado1.Visible = False
'            ucFacturacionProductos.Visible = True
'            ucFacturacionProductos.EstadosFacturacion = "5"    'Autorizados a devolver
'            ucFacturacionProductos.TiposFinanciamiento = "1,5,9"
'            ucFacturacionProductos.IdOrden = IdOrden
'            ucFacturacionProductos.CargaProductosPorIdOrden
'        Case sghopcionespago.sghAnulacion
'            UcFacturacionContado1.Visible = False
'            ucFacturacionProductos.Visible = True
'            'ucFacturacionProductos.EstadosFacturacion = "4"    'Pagados
'            ucFacturacionProductos.EstadosFacturacion = IIf(mo_DOComprobantePago.IdTipoPago = 2, 6, 4)
'            ucFacturacionProductos.TiposFinanciamiento = "1,2,3,5,9"
'            ucFacturacionProductos.IdOrden = IdOrden
'            ucFacturacionProductos.CargaProductosPorIdOrden
'        End Select
'        Set oDoPreVenta = Nothing
'        Set oPreVenta = Nothing
'        Set oConexion = Nothing
'End Sub

'debb-18/02/2011
'Private Sub txtNroHistoria_LostFocus()
'    If txtNroHistoria.Text <> "" Then
'       If Len(txtNroHistoria.Text) > 9 Or mo_Teclado.TextoEsSoloNumeros(txtNroHistoria.Text) = False Then
'          MsgBox "La longitud no puede pasar de 9", vbInformation, "caja"
'          txtNroHistoria.Text = ""
'          Exit Sub
'       End If
'       txtNroCuenta.Text = ""
'       LimpiarFormulario
'       If cmbIdTipoGenHistoriaClinica.Text <> "" Then
'            Dim oRsTmp As New Recordset
'            Dim lnIdPacienteHallado As Long
'            Dim oConexion As New Connection
'            oConexion.Open sighEntidades.CadenaConexion
'            oConexion.CursorLocation = adUseClient
'
'            Set oRsTmp = mo_AdminAdmision.PacientesXnroHistoriaTipoNumeracion(Val(txtNroHistoria.Text), Val(mo_cmbIdTipoGenHistoriaClinica.BoundText), oConexion)
'            If oRsTmp.RecordCount > 0 Then
'               lnIdPacienteHallado = oRsTmp.Fields!idPaciente
'               If Not IsNull(oRsTmp.Fields!NroDocumento) And Not IsNull(oRsTmp.Fields!IdDocIdentidad) Then
'                  If oRsTmp.Fields!IdDocIdentidad = 1 Then
'                     txtDni.Text = oRsTmp.Fields!NroDocumento
'                  End If
'               End If
'               oRsTmp.Close
'               Set oRsTmp = mo_ReglasFarmacia.FacturacionCuentasAtencionSeleccionarPorIdPaciente(lnIdPacienteHallado, oConexion)
'               If oRsTmp.RecordCount > 0 Then
'                  txtNroCuenta.Text = oRsTmp.Fields!idCuentaAtencion
'               Else
'                  MsgBox "Ese Nro de HISTORIA existe, pero no tiene Nro Cuenta", vbInformation, "Caja"
'               End If
'               oRsTmp.Close
'               Set oRsTmp = Nothing
'               Else
'               MsgBox "Ese Nro de HISTORIA no existe", vbInformation, "Caja"
'               oRsTmp.Close
'               Set oRsTmp = Nothing
'            End If
'            oConexion.Close
'            Set oConexion = Nothing
'        Else
'            MsgBox "Elija 'Tipo Historia'", vbInformation, "Caja"
'        End If
'    End If
'End Sub

Sub LeerServiciosPorTipoDePago()

     Select Case mi_OpcionINO
'        Case sghopcionespago.sghNuevoPagoConHistoria
'
'        Case sghopcionespago.sghNuevoPagoSinHistoria
'
'        Case sghopcionespago.sghPagarOrdenExistente
'           CargarDatosServiciosALosControlesPorIdOrden
'            txtRazonSocial.SetFocus
'        Case sghopcionespago.sghPagarCuentaExistente
        
        Case sghopcionespago.sghDevolucion
            
            CargarDatosALosControlesPorNroSerieBoleta
'        Case sghopcionespago.sghAnulacion
'
'            CargarDatosALosControlesPorNroSerieBoleta
'             btnAceptar.SetFocus
'        Case sghopcionespago.sghReimprimirComprobante
'            CargarDatosALosControlesPorNroSerieBoleta
        
        
        
        '/****************************************************/
        '/***********************INO**************************/
        '/****************************************************/
        
        
        Case sghopcionespago.sghDevolucionINO
        Dim buscarIDAtencion As New Recordset
        Set buscarIDAtencion = mo_AdminCaja.CajaComprobantesPagoSeleccionarPorNroSerieNroDocumento(txtNserieB.Text, txtNdocumentoB.Text)
        
        If buscarIDAtencion.RecordCount > 0 Then
            Dim orst2 As New Recordset
            Dim orstOft As New Recordset
            Set orst2 = mo_AdminCaja.BuscarTriaje(buscarIDAtencion.Fields!idCuentaAtencion)
            'Set orstOft = mo_AtencionesOftalmologicas.BuscarTriajeOftalmologico(buscarIDAtencion.Fields!idCuentaAtencion)
            If orst2.RecordCount = 1 Then
                MsgBox "La cita ya ha sido atendida!", vbInformation, "Caja"
'            ElseIf orstOft.RecordCount = 1 Then
'                MsgBox "La cita ya ha sido atendida (triaje oftalmológico)!", vbInformation, "Caja"
            Else
                CargarDatosALosControlesPorNroSerieBoleta
                btnAceptar.SetFocus
            End If
        Else
            MsgBox "El Documento " & txtNroSerie.Text & " - " & txtNroDocumento.Text & " NO EXISTE", vbInformation, "Caja"
        End If
  
        Case sghopcionespago.sghAnularDevolucionINO
'            Dim orst As New Recordset
'            Set orst = mo_AdminCaja.CajaBuscarDevolucion(txtNserieB.Text, txtNdocumentoB.Text)
'
'            If orst.RecordCount = 0 Then
'                MsgBox "El comprobante de pago que esta buscando no ha sido devuelto", vbInformation, "Caja"
'            Else
'                CargarDatosALosControlesPorNroSerieBoleta
'                btnAceptar.SetFocus
'            End If
        
        '/****************************************************/
        '/****************************************************/
        '/****************************************************/
    End Select
    
End Sub

Sub LeerBienesPorTipoDePago()

     Select Case mi_OpcionINO
'        Case sghopcionespago.sghNuevoPagoConHistoria
'
'        Case sghopcionespago.sghNuevoPagoSinHistoria
'
'        Case sghopcionespago.sghPagarOrdenExistenteF
'           CargarDatosBienesALosControlesPorIdOrden
'           txtRazonSocial.SetFocus
'        Case sghopcionespago.sghPagarCuentaExistente
        
        Case sghopcionespago.sghDevolucion
            CargarDatosALosControlesPorNroSerieBoleta
'        Case sghopcionespago.sghAnulacion
'
'            CargarDatosALosControlesPorNroSerieBoleta
'            btnAceptar.SetFocus
'        Case sghopcionespago.sghReimprimirComprobante
'            CargarDatosALosControlesPorNroSerieBoleta
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
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    If Val(txtNserieB.Text) = 0 Or Val(txtNdocumentoB.Text) = 0 Then
       Exit Sub
    End If
    Set rsBuscaBoleta = oReglasCaja.CajaComprobantesPagoSeleccionarPorNroSerieNroDocumento(txtNserieB.Text, txtNdocumentoB.Text)
    If rsBuscaBoleta.RecordCount = 0 Then
       MsgBox "El Documento " & txtNserieB.Text & " - " & txtNdocumentoB.Text & " NO EXISTE", vbInformation, "Caja"
       Exit Sub
    End If
    If BoletaTieneRegistradoLaboratorioImagenes(rsBuscaBoleta.Fields!IdComprobantePago) = True Then
       Exit Sub
    End If
    txtNroSerie.Text = Trim(txtNserieB.Text)
    txtNroDocumento.Text = Trim(txtNdocumentoB.Text)
    lbBoletaDeServicios = IIf(rsBuscaBoleta.Fields!IdTipoOrden = 1, True, False)
    If lbBoletaDeServicios Then
       Set rsBuscaBoleta = oReglasCaja.CajaComprobantePagoServiciosPorNroSerieNroDocumento(txtNroSerie.Text, txtNroDocumento.Text)
    Else
       Set rsBuscaBoleta = oReglasCaja.CajaComprobantePagoProductosPorNroSerieNroDocumento(txtNroSerie.Text, txtNroDocumento.Text)
    End If
    If rsBuscaBoleta.RecordCount = 0 Then
       MsgBox "El Documento " & txtNroSerie.Text & " - " & txtNroDocumento.Text & " NO EXISTE", vbInformation, "Caja"
    ElseIf rsBuscaBoleta.Fields!IdEstadoComprobante = 9 Then
       MsgBox "El Documento " & txtNroSerie.Text & " - " & txtNroDocumento.Text & " YA ESTA ANULADO", vbInformation, "Caja"
    '/**************************************************/
    '/*********************INO**************************/
    '/**************************************************/
    ElseIf rsBuscaBoleta.Fields!IdEstadoComprobante = 6 And optDevoluciones.Value = True Then
           MsgBox "El Documento " & txtNroSerie.Text & " - " & txtNroDocumento.Text & " YA HA SIDO DEVUELTO", vbInformation, "Caja"
    '/**************************************************/
    '/*********************INO**************************/
    '/**************************************************/
    ElseIf rsBuscaBoleta.Fields!IdEstadoComprobante = 1 Then
       MsgBox "La orden aun no ha sido PAGADA, solo se puede realizar anulaciones de ordenes PAGADAS.", vbInformation, "Caja"
'    ElseIf TienePaqueteDespachado(rsBuscaBoleta.Fields!IdComprobantePago) = True Then
'       MsgBox "No se podrá Anular porque ya se ha Atendido Parte/Total del PAQUETE", vbInformation, "Caja"
    Else
       'Carga Cabecera
       LimpiarOpciones
       txtNroHistoria.Text = Trim(IIf(IsNull(rsBuscaBoleta.Fields!NroHistoriaClinica), "", rsBuscaBoleta.Fields!NroHistoriaClinica))
       ml_IdPaciente = IIf(IsNull(rsBuscaBoleta.Fields!idPaciente), 0, rsBuscaBoleta.Fields!idPaciente)
       txtRazonSocial.Text = Trim(IIf(IsNull(rsBuscaBoleta.Fields!RazonSocial), "", rsBuscaBoleta.Fields!RazonSocial))
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
       If mi_OpcionINO = sghopcionespago.sghDevolucion Then
            ucFacturacionProductos.ActualizaDevolucionAutorizada rsBuscaBoleta
            Dim lnTotal As Double
            rsBuscaBoleta.MoveFirst
            Do While Not rsBuscaBoleta.EOF
               If Not IsNull(rsBuscaBoleta.Fields!cantidadDev) Then
               lnTotal = lnTotal + rsBuscaBoleta.Fields!PrecioUnitario * rsBuscaBoleta.Fields!cantidadDev
               End If
               rsBuscaBoleta.MoveNext
            Loop
'            txtPagoACuenta.Text = 0
            txtTotal.Text = Format(lnTotal, "#######.#0")
            txtVuelto.Text = Format("0", "#######.#0")
       Else
             
            rsBuscaBoleta.MoveFirst
'            txtPagoACuenta.Text = Format(rsBuscaBoleta.Fields!Adelantos, "#######.#0")
'            txtExonerado.Text = Format(rsBuscaBoleta.Fields!Exoneraciones, "#######.#0")
            txtTotal.Text = Format(rsBuscaBoleta.Fields!TotalBoleta, "#######.#0")
            txtVuelto.Text = Format(rsBuscaBoleta.Fields!TotalBoleta, "#######.#0")
       End If

    End If
    Set oReglasCaja = Nothing
    Set rsBuscaBoleta = Nothing
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
'
'Sub ConfigurarFechaIngreso()
'
'    mo_cmbFechaIngreso.ListField = "DescripcionLarga"
'    mo_cmbFechaIngreso.BoundColumn = "IdCuentaAtencion"
'
'End Sub
'
'Sub ConfigurarPuntosDeCarga()
'
'    mo_cmbIdPuntoCarga.ListField = "Descripcion"
'    mo_cmbIdPuntoCarga.BoundColumn = "IdPuntoCarga"
'    Set mo_cmbIdPuntoCarga.RowSource = mo_reglasComunes.SeleccionarPuntosDeCarga()
'
'End Sub

'Sub ConfigurarTiposHistoriaClinica()
'
'        mo_cmbIdTipoGenHistoriaClinica.BoundColumn = "IdTipoNumeracion"
'        mo_cmbIdTipoGenHistoriaClinica.ListField = "DescripcionLarga"
'        Set mo_cmbIdTipoGenHistoriaClinica.RowSource = mo_AdminArchivoClinico.TiposGeneracionHistoriasSeleccionarTodos()
'
'End Sub
'
'Sub ConfigurarTipoComprobante()
'
'    mo_cmbIdTipoComprobante.BoundColumn = "IdTipoComprobante"
'    mo_cmbIdTipoComprobante.ListField = "Descripcion"
'    Set mo_cmbIdTipoComprobante.RowSource = mo_AdminCaja.TiposComprobanteSeleccionarTodos()
'
'End Sub

                
'debb-17/02/2011
Sub CargaDatosDeTotalesDeLaCuenta(lIdCuentaAtencion As Long)
    Dim lnTotalDctosPorAdelantos As Double
    Dim lnTotalPagarFarmacia As Double
    Dim lnPagosXdevoluciones As Double
    Dim oConexion As New Connection
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    lnTotalDctosPorAdelantos = mo_AdminCaja.RetornaTotalDescuentosPorAdelantosSegunCuenta(lIdCuentaAtencion, oConexion)
    lnPagosXdevoluciones = mo_ReglasFacturacion.RetornaImporteDePagosXdevolucionesPorNroCuenta(ml_idCuentaAtencion, oConexion)
    lnTotalDctosPorAdelantos = lnTotalDctosPorAdelantos - lnPagosXdevoluciones
    If (lnTotalDctosPorAdelantos) > 0 Then
       lnTotalPagarFarmacia = mo_ReglasFacturacion.RetornaTotalPagosFarmaciaPendientesPorNroCuentadebb(lIdCuentaAtencion)
       If tabFactProductosPorCuenta.Tab = 0 Then
          '****SERVICIOS: disminuir consumo de FARMACIA del Adelanto
'          txtPagoACuenta.Text = lnTotalDctosPorAdelantos - lnTotalPagarFarmacia
       Else
          '****FARMACIA:
          lnTotalGrid = lnTotalPagarFarmacia
          If (lnTotalDctosPorAdelantos) > lnTotalPagarFarmacia Then
'             txtPagoACuenta.Text = lnTotalPagarFarmacia
          Else
'             txtPagoACuenta.Text = lnTotalDctosPorAdelantos
          End If
       End If
    End If
'    ActualizaTotalApagar
'    If txtTotal.Text < 0 Then
''       txtPagoACuenta.Text = CCur(txtPagoACuenta.Text) + CCur(txtTotal.Text)
'       ActualizaTotalApagar
'    End If
End Sub
            
'Sub ActualizaTotalApagar()
'    If txtExonerado.Text = "" Then
'       txtExonerado.Text = "0"
'    End If
'    If txtPagoACuenta.Text = "" Then
'       txtPagoACuenta.Text = "0"
'    End If
'    txtTotal.Text = lnTotalGrid - CCur(txtExonerado.Text) - CCur(txtPagoACuenta.Text)
'    txtTotal.Text = sighEntidades.DevuelveNumeroRedondeado(CCur(txtTotal.Text))   'debb-mayo2014
'    txtEfectivo.Text = txtTotal.Text
'End Sub


Sub ConfiguraPermisos()
    'PERMISOS
    Dim oRsPermisos As New Recordset
    Set oRsPermisos = mo_ReglasSeguridad.UsuariosRolesSeleccionarPermisosTodos(ml_idUsuario)
    UserControl.tabGestionCaja.TabVisible(1) = False
    lbTienePermisoSoloParaBoletaFarmacia = False
    lbTienePermisoReimprimeBoleta = False
    lbTienePermisoExonerarPacExterno = False
    lbTienePermisoSoloParaBoletaServicio = False
    If oRsPermisos.RecordCount > 0 Then
       Do While Not oRsPermisos.EOF
          Select Case oRsPermisos.Fields!IdPermiso
          Case 206    'Caja - Ver TAB 'Registro de Comprobante'
               UserControl.tabGestionCaja.TabVisible(1) = True
          Case 366    'Caja - Sólo Emite Boleta p' Pre-Venta Farmacia
               lbTienePermisoSoloParaBoletaFarmacia = True
          Case 367    'Caja - Reimprime Boleta
               lbTienePermisoReimprimeBoleta = True
          Case 368    'Caja - Pacientes Externos - permite ingresar exoneracion
               lbTienePermisoExonerarPacExterno = True
          Case 369    'Caja - Sólo Emite Boletas de Servicios - sólo CPT
               lbTienePermisoSoloParaBoletaServicio = True
          End Select
          oRsPermisos.MoveNext
       Loop
    End If
    Set oRsPermisos = Nothing
End Sub




'/******************************************************/
'/***********************INO****************************/
'/******************************************************/


Private Sub optDevoluciones_Click()
 'mi_OpcionINO = sghOpcionesPago.sghAnulacion
    mi_OpcionINO = sghopcionespago.sghDevolucionINO
    fraOpciones.Visible = True
    
    LimpiarOpciones
    LimpiarFormulario
    mo_Formulario.HabilitarDeshabilitar txtNserieB, True
    mo_Formulario.HabilitarDeshabilitar txtNdocumentoB, True
      
    tabFactProductosPorCuenta.Visible = False
    lbEsDevolucion = False
'    txtNreceta.Enabled = False: cmbBuscaReceta.Enabled = False
    InabilitaTotales
    txtNserieB.SetFocus
  
End Sub



Private Sub optAnularDevoluciones_Click()
    mi_OpcionINO = sghopcionespago.sghAnularDevolucionINO
    
    fraOpciones.Visible = True
    
    LimpiarOpciones
    LimpiarFormulario
    mo_Formulario.HabilitarDeshabilitar txtNserieB, True
    mo_Formulario.HabilitarDeshabilitar txtNdocumentoB, True
       
    tabFactProductosPorCuenta.Visible = False
    lbEsDevolucion = False
'    txtNreceta.Enabled = False: cmbBuscaReceta.Enabled = False
    
    InabilitaTotales
    txtNserieB.SetFocus
End Sub

Function DevolucionINO() As Boolean
    mo_DOComprobantePago.IdUsuarioAuditoria = ml_idUsuario
    If lbBoletaDeServicios Then
     CargarDatosDevolucion
     DevolucionINO = mo_AdminCaja.CajaDevolucionesAgregar(mo_DOCajaDevolucion)
     'CambioEstadoComprobanteINO
     ms_MensajeError = mo_AdminCaja.MensajeError
    End If
End Function

Function AnularDevolucionINO() As Boolean

    mo_DOComprobantePago.IdUsuarioAuditoria = ml_idUsuario
    'mo_DOCajaDevolucion.idDevolucion = 3
    mo_DOCajaDevolucion.IdComprobantePago = mo_DOComprobantePago.IdComprobantePago
    idComprobantePagoINO = mo_DOCajaDevolucion.IdComprobantePago
     AnularDevolucionINO = mo_AdminCaja.CajaDevolucionesEliminar(mo_DOCajaDevolucion)
     ms_MensajeError = mo_AdminCaja.MensajeError

End Function


Function Devolucion() As Boolean
        Select Case ml_TipoProductoINO
    Case sghServicio
        Devolucion = mo_AdminCaja.CajaComprobantePagodevolucionOrdenServicio(mo_DOComprobantePago.IdComprobantePago, mo_DOComprobantePagoDevolucion, ml_idUsuario)
    Case sghbien
        Devolucion = mo_AdminCaja.CajaComprobantePagoDevolucionOrdenBienInsumo(mo_DOComprobantePago.IdComprobantePago, mo_DOComprobantePagoDevolucion, ml_idUsuario)
    End Select
End Function


 Sub CargarDatosDevolucion()
    mo_DOCajaDevolucion.IdComprobantePago = mo_DOComprobantePago.IdComprobantePago
    idComprobantePagoINO = mo_DOComprobantePago.IdComprobantePago
    mo_DOCajaDevolucion.montoDevuelto = txtTotal.Text
    mo_DOCajaDevolucion.montoTotal = txtTotal.Text
    mo_DOCajaDevolucion.fechaDevolucion = lcBuscaParametro.RetornaFechaHoraServidorSQL
    mo_DOCajaDevolucion.idUsuario = ml_idUsuario
   ' mo_DOCajaDevolucion.mMotivo = mo_CajaObservacion.mMotivo
        
 End Sub
''
'' Function CambioEstadoComprobanteINO() As Boolean
''    Dim oConexion As New Connection
''    oConexion.Open sighEntidades.CadenaConexion
''    oConexion.CursorLocation = adUseClient
''    mo_DOComprobantePago.IdUsuarioAuditoria = ml_idUsuario
''   ' If lblboletadeservicios Then
''        CambioEstadoComprobanteINO = mo_AdminCaja.ComprobantePagoSeleccionarPorId(idComprobantePagoINO, oConexion)
''        ms_MensajeError = mo_AdminCaja.MensajeError
''    'End If
'' End Function
''
''
''Sub CargarDatosEstadoComprobanteINO()
''    mo_DOCajaComprobante.IdEstadoComprobante = sghEstadosComprobante.sighEstadosComprobanteDevuelto
''End Sub

'/**************************************************/
'/**************************************************/






Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property


Private Sub btnBuscar_Click()
    RealizarBusqueda
End Sub




Private Sub cmdSinApellidoMaterno_Click()
    txtApellidoMaterno.Text = wxSinApellido
End Sub

Private Sub cmdSinApellidoPaterno_Click()
    txtApellidoPaterno.Text = wxSinApellido
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub MuestraTabEmisionDocumentos(lbVisible As Boolean)
        UserControl.tabGestionCaja.TabVisible(1) = lbVisible
End Sub

'
'Private Sub Form_Load()
'    mb_lbEstoyEnTabServicio = False
'    txtFecha1.Text = Date
'    txtFecha2.Text = Date
''    If ml_TipoProducto <> sghAmbos Then
''       Me.tabFarmServ.TabVisible(1) = False
''       Me.tabFarmServ.TabCaption(0) = ""
''    End If
'   ' BusquedaDatos
'  mo_Apariencia.ConfigurarFilasBiColores grdDevolucion, sighEntidades.GrillaConFilasBicolor
''    mo_Apariencia.ConfigurarFilasBiColores Me.grdTab1, sighEntidades.GrillaConFilasBicolor
'End Sub
Public Sub RealizarBusqueda()
'Dim oDoMedico As New DOMedico
'Dim oDOEmpleado As New dOEmpleado
'Dim oDoDevolucion As New DOCajaDevoluciones
'Dim lComprobante As String

'oDoDevolucion.idDevolucion = txtComprobante.Text

'        oDOEmpleado.ApellidoPaterno = UserControl.txtApellidoPaterno
'        oDOEmpleado.ApellidoMaterno = UserControl.txtApellidoMaterno
'        oDOEmpleado.Nombres = UserControl.txtNombres
'        oDOEmpleado.CodigoPlanilla = UserControl.txtCodigoPlanilla
'        If lcBuscaParametro.SeleccionaFilaParametro(264) = "S" And (ml_IdTipoServicio = 2 Or ml_IdTipoServicio = 4) Then
'           'Para que funcione en Hospitalización falta poner  .....or ml_idTipoServicio = 3....
         If txtComprobante.Text = "0" Then
           Set grdDevolucion.DataSource = mo_AdminCaja.BuscarDevoluciones(txtComprobante)
         ElseIf txtComprobante.Text = "" And txtApellidoPaterno.Text = "" And txtApellidoMaterno.Text = "" Then
           Set grdDevolucion.DataSource = mo_AdminCaja.CajaDevolucionesPorFechas(CDate(txtFecha1.Text), CDate(txtFecha2.Text))
         ElseIf txtComprobante.Text = "" And txtApellidoPaterno.Text = "" Then
           Set grdDevolucion.DataSource = mo_AdminCaja.BuscarDevolucionesApellido2(txtApellidoMaterno)
         'ElseIf txtComprobante.Text = "" Then
         ElseIf txtComprobante.Text = "" Then
           Set grdDevolucion.DataSource = mo_AdminCaja.BuscarDevolucionesApellido(txtApellidoPaterno)
         
     End If
'        If mo_AdminProgramacionMedica.MensajeError <> "" Then
'            MsgBox "Error leyendo datos" + Chr(13) + mo_AdminProgramacionMedica.MensajeError, vbInformation, "Profesional de la Salud"
'        End If
        
        mo_Apariencia.ConfigurarFilasBiColores grdDevolucion, sighentidades.GrillaConFilasBicolor

End Sub



'Private Sub grdDevolucion_DblClick()
'    Dim oRsTmp As New ADODB.Recordset
'    Set oRsTmp = grdDevolucion.DataSource
'    If oRsTmp.RecordCount > 0 Then
'        If ml_TipoProducto = sghbien Then
'            ml_idOrdenSeleccionado = IIf(IsNull(oRsTmp.Fields!Nro_Preventa), 0, oRsTmp.Fields!Nro_Preventa)
'        ElseIf ml_TipoProducto = sghServicio Then
'            ml_idOrdenSeleccionado = oRsTmp.Fields!IdOrdenPago
'        Else
'            ml_idOrdenSeleccionado = IIf(IsNull(oRsTmp.Fields!Nro_Preventa), 0, oRsTmp.Fields!Nro_Preventa)
'        End If
'        mi_BotonPresionado = sghAceptar
'        Me.Visible = False
'    End If
'End Sub

'Private Sub grdDevolucion_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
'    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
'End Sub

'Private Sub grdDevolucion_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
'    If KeyAscii = 13 Then
'        grdDevolucion_DblClick
'    End If
'End Sub


'Private Sub grdTab1_DblClick()
'    Dim oRsTmp As New ADODB.Recordset
'    Set oRsTmp = Me.grdTab1.DataSource
'    If oRsTmp.RecordCount > 0 Then
'        ml_idOrdenSeleccionado = oRsTmp.Fields!IdOrdenPago
'        mi_BotonPresionado = sghAceptar
'        Me.Visible = False
'    End If
'End Sub



Private Sub grdTab1_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
End Sub

Private Sub grdTab1_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = 13 Then
'        grdTab1_DblClick
    End If

End Sub


Private Sub txtComprobante_LostFocus()
    If Len(txtComprobante.Text) > 0 Then
       btnBuscar_Click
    End If
End Sub



Private Sub txtFecha1_LostFocus()
    If Not IsDate(txtFecha1.Text) Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        btnLimpiar_Click
        Exit Sub
    End If
End Sub

Private Sub txtFecha2_LostFocus()
    If Not IsDate(txtFecha2.Text) Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        btnLimpiar_Click
        Exit Sub
    End If
End Sub

Private Sub btnLimpiar_Click()
   
     txtFecha1.Text = Date & " 00:01"
     txtFecha2.Text = Date & " 23:59"
    txtApellidoPaterno.Text = ""
    txtApellidoMaterno.Text = ""
    txtComprobante.Text = ""
End Sub




