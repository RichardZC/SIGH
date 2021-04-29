VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form AdmisionCEDetalle 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9675
   ClientLeft      =   1830
   ClientTop       =   -105
   ClientWidth     =   11910
   ControlBox      =   0   'False
   Icon            =   "AdmisionDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9675
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SISGalenPlus.ucMensajeParpadeando ucMensajeParpadeando2 
      Height          =   285
      Left            =   8910
      TabIndex        =   101
      Top             =   60
      Visible         =   0   'False
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   503
   End
   Begin SISGalenPlus.UcSISafiliacion UcSISafiliacion1 
      Height          =   600
      Left            =   2640
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1270
   End
   Begin UltraGrid.SSUltraGrid grdPacientesEncontrados 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   661
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
      Caption         =   "Lista de pacientes encontrados"
   End
   Begin TabDlg.SSTab tabAdmision 
      Height          =   6915
      Left            =   0
      TabIndex        =   14
      Top             =   1920
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   12197
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
      TabCaption(0)   =   "Paciente (F10)"
      TabPicture(0)   =   "AdmisionDetalle.frx":0CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ucPacientesDetalle1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Cita (F11)"
      TabPicture(1)   =   "AdmisionDetalle.frx":0CE6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "TabIngreso"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin SISGalenPlus.ucPacientesDetalle ucPacientesDetalle1 
         Height          =   6495
         Left            =   -74910
         TabIndex        =   9
         Top             =   330
         Width           =   11745
         _ExtentX        =   20823
         _ExtentY        =   11351
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   6405
         Left            =   -74850
         TabIndex        =   18
         Top             =   420
         Width           =   11745
         _ExtentX        =   20717
         _ExtentY        =   11298
         _Version        =   393216
         Tabs            =   1
         TabHeight       =   520
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Diagnósticos"
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).ControlCount=   0
      End
      Begin TabDlg.SSTab TabIngreso 
         Height          =   6465
         Left            =   90
         TabIndex        =   26
         Top             =   390
         Width           =   11985
         _ExtentX        =   21140
         _ExtentY        =   11404
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         ForeColor       =   13653559
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "2.1 Atención"
         TabPicture(0)   =   "AdmisionDetalle.frx":0D02
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "UcEpisodioClinico1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "fraDatosCita"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "fraDatosReferenciaOrigen"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Frame5"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "FraGeneraCita"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "2.2 Citas para otros días"
         TabPicture(1)   =   "AdmisionDetalle.frx":0D1E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "UcPacientesSunasa1"
         Tab(1).Control(1)=   "ucCitasLista11"
         Tab(1).Control(2)=   "Label9"
         Tab(1).ControlCount=   3
         Begin VB.Frame FraGeneraCita 
            Caption         =   "Forma que se genera la CITA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   5940
            TabIndex        =   122
            Top             =   5385
            Width           =   5775
            Begin Threed.SSOption optCNorma 
               Height          =   270
               Left            =   195
               TabIndex        =   123
               Top             =   255
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   476
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
               Caption         =   "Normal"
               Value           =   -1
            End
            Begin Threed.SSOption optCweb 
               Height          =   270
               Left            =   2565
               TabIndex        =   124
               Top             =   255
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   476
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
               Caption         =   "Web"
            End
            Begin Threed.SSOption optCtelefono 
               Height          =   270
               Left            =   4560
               TabIndex        =   125
               Top             =   255
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   476
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
               Caption         =   "Teléfono"
            End
         End
         Begin SISGalenPlus.UcPacientesSunasa UcPacientesSunasa1 
            Height          =   375
            Left            =   -65340
            TabIndex        =   91
            Top             =   45
            Visible         =   0   'False
            Width           =   1965
            _ExtentX        =   20294
            _ExtentY        =   10451
         End
         Begin VB.Frame Frame5 
            Height          =   4995
            Left            =   5910
            TabIndex        =   80
            Top             =   330
            Width           =   5805
            Begin VB.CommandButton cmdAgregaHistorico 
               Caption         =   "..."
               Height          =   315
               Left            =   3615
               TabIndex        =   126
               Top             =   2730
               Width           =   135
            End
            Begin SISGalenPlus.ucEPS ucEPS1 
               Height          =   345
               Left            =   180
               TabIndex        =   121
               Top             =   1905
               Width           =   5520
               _ExtentX        =   9737
               _ExtentY        =   609
            End
            Begin VB.ComboBox cmbFormaPago 
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
               Left            =   1920
               Style           =   2  'Dropdown List
               TabIndex        =   98
               Top             =   660
               Width           =   3840
            End
            Begin VB.ComboBox cmbFuenteFinanciamiento 
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
               Left            =   1920
               Style           =   2  'Dropdown List
               TabIndex        =   97
               Top             =   240
               Width           =   3840
            End
            Begin VB.TextBox txtNroOrdenPago 
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
               Height          =   315
               Left            =   1905
               TabIndex        =   84
               Top             =   2730
               Width           =   1665
            End
            Begin VB.TextBox txtNroCuenta 
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
               Height          =   315
               Left            =   1905
               TabIndex        =   83
               Top             =   2340
               Width           =   1665
            End
            Begin VB.TextBox txtNserie 
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
               Left            =   1920
               MaxLength       =   4
               TabIndex        =   82
               Top             =   1080
               Visible         =   0   'False
               Width           =   555
            End
            Begin VB.TextBox txtNboleta 
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
               Left            =   2460
               MaxLength       =   30
               TabIndex        =   81
               Top             =   1080
               Visible         =   0   'False
               Width           =   1125
            End
            Begin SISGalenPlus.ucSISfuaCodPrestacion ucSISfuaCodPrestacion1 
               Height          =   345
               Left            =   180
               TabIndex        =   100
               Top             =   1500
               Width           =   5505
               _ExtentX        =   9710
               _ExtentY        =   609
            End
            Begin VB.Label lblNroAtencion 
               AutoSize        =   -1  'True
               Caption         =   "Nº Atencion"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   240
               Left            =   165
               TabIndex        =   93
               Top             =   3210
               Width           =   1005
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Producto/Plan"
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
               Left            =   180
               TabIndex        =   90
               Top             =   690
               Width           =   1185
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fte.Financiam/IAFA"
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
               Left            =   180
               TabIndex        =   89
               Top             =   270
               Width           =   1680
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nº Orden de Pago"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   240
               Left            =   180
               TabIndex        =   88
               Top             =   2790
               Width           =   1545
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nº Cuenta"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   240
               Left            =   180
               TabIndex        =   87
               Top             =   2370
               Width           =   870
            End
            Begin VB.Label lblEstadoCta 
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
               Left            =   3690
               TabIndex        =   86
               Top             =   2370
               Width           =   180
            End
            Begin VB.Label lblBoleta 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "N° Boleta (Paquete)"
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
               Left            =   180
               TabIndex        =   85
               Top             =   1170
               Visible         =   0   'False
               Width           =   1680
            End
         End
         Begin VB.Frame fraDatosReferenciaOrigen 
            Caption         =   " Origen de  referencia "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   2145
            Left            =   120
            TabIndex        =   74
            Top             =   3900
            Width           =   5775
            Begin VB.CommandButton btnDxReferencia 
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
               Left            =   2670
               Picture         =   "AdmisionDetalle.frx":0D3A
               Style           =   1  'Graphical
               TabIndex        =   118
               Top             =   1365
               Width           =   330
            End
            Begin VB.CommandButton btnBuscarEstablecimiento 
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
               Left            =   2685
               Picture         =   "AdmisionDetalle.frx":12C4
               Style           =   1  'Graphical
               TabIndex        =   117
               Top             =   660
               Width           =   330
            End
            Begin VB.TextBox txtMedicoRef 
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
               Left            =   1635
               MaxLength       =   8
               TabIndex        =   115
               ToolTipText     =   "Busca por COLEGIATURA"
               Top             =   1725
               Width           =   1000
            End
            Begin VB.ComboBox cmbMedicoRef 
               BackColor       =   &H00FFFFFF&
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
               Height          =   330
               Left            =   2670
               TabIndex        =   114
               Top             =   1710
               Width           =   2985
            End
            Begin VB.TextBox lblDxReferencia1 
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
               Left            =   3030
               TabIndex        =   111
               Top             =   1350
               Width           =   2610
            End
            Begin VB.TextBox txtDxReferencia 
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
               Left            =   1650
               MaxLength       =   8
               TabIndex        =   104
               Top             =   1350
               Width           =   1000
            End
            Begin VB.TextBox txtReferenciaO 
               Height          =   315
               Left            =   4635
               MaxLength       =   20
               TabIndex        =   102
               Top             =   285
               Width           =   1020
            End
            Begin VB.TextBox txtIdEstablecimientoOrigen 
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
               Left            =   1650
               MaxLength       =   8
               TabIndex        =   77
               Top             =   660
               Width           =   1000
            End
            Begin VB.TextBox txtNombreOrigenReferencia 
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
               Left            =   3030
               TabIndex        =   76
               Top             =   660
               Width           =   2610
            End
            Begin VB.ComboBox cmbIdTipoReferenciaOrigen 
               BackColor       =   &H00FFFFFF&
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
               Height          =   330
               Left            =   1650
               TabIndex        =   75
               Top             =   300
               Width           =   1695
            End
            Begin PVCOMBOLibCtl.PVComboBox cmbServicioReferenciaO 
               Height          =   330
               Left            =   1650
               TabIndex        =   103
               Top             =   1005
               Width           =   4020
               _Version        =   524288
               _cx             =   7091
               _cy             =   582
               Appearance      =   1
               Enabled         =   -1  'True
               BackColor       =   16777215
               ForeColor       =   0
               Locked          =   0   'False
               Style           =   0
               Sorted          =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ShowPictures    =   0   'False
               ColumnHeaders   =   -1  'True
               PrimaryColumn   =   1
               VisibleItems    =   10
               ColumnHeaderHeight=   20
               ListMember      =   ""
               ColumnHeaderForeColor=   0
               ColumnHeaderBackColor=   13160660
               SelectedForeColor=   16777215
               SelectedBackColor=   6956042
               AlternateBackColor=   16777215
               ItemLabelStyle  =   1
               ItemLabelType   =   0
               ItemLabelWidth  =   40
               ItemLabelForeColor=   0
               ItemLabelBackColor=   13160660
               ColumnHeaderStyle=   1
               VerticalGridLines=   -1  'True
               HorizontalGridLines=   -1  'True
               ColumnResize    =   0   'False
               ItemLabelResize =   0   'False
               AllowDBAutoConfig=   0   'False
               GridLineColor   =   13421772
               List            =   ""
               NullString      =   "[NULL]"
               DropShadow      =   -1  'True
               Text            =   ""
               SortOnColumnHeaderClick=   0   'False
               DropEffect      =   1
               ColumnCount     =   2
               Column0.Heading =   "Código"
               Column0.Width   =   60
               Column0.Alignment=   0
               Column0.Hidden  =   0   'False
               Column0.Name    =   "codigo"
               Column0.Format  =   ""
               Column0.Bound   =   -1  'True
               Column0.Locked  =   0   'False
               Column0.HeaderAlignment=   0
               Column1.Heading =   "Descripción"
               Column1.Width   =   200
               Column1.Alignment=   0
               Column1.Hidden  =   0   'False
               Column1.Name    =   "descripcion"
               Column1.Format  =   ""
               Column1.Bound   =   -1  'True
               Column1.Locked  =   0   'False
               Column1.HeaderAlignment=   0
               SortKey1.Column =   -1
               SortKey1.Ascending=   -1  'True
               SortKey1.CaseInsensitive=   -1  'True
               SortKey2.Column =   -1
               SortKey2.Ascending=   -1  'True
               SortKey2.CaseInsensitive=   -1  'True
               SortKey3.Column =   -1
               SortKey3.Ascending=   -1  'True
               SortKey3.CaseInsensitive=   -1  'True
               BoundColumn     =   ""
               Border          =   -1  'True
               VertAlign       =   1
               Format          =   ""
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Médico referencia"
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
               Left            =   105
               TabIndex        =   116
               Top             =   1770
               Width           =   1440
            End
            Begin VB.Label lblDxreferencia 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dx referencia"
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
               TabIndex        =   112
               Top             =   1395
               Width           =   1080
            End
            Begin VB.Label lblServicioReferencia 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Servicio referencia"
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
               TabIndex        =   110
               Top             =   1035
               Width           =   1485
            End
            Begin VB.Label lblReferenciaO 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "N° Referencia"
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
               Left            =   3495
               TabIndex        =   105
               Top             =   315
               Width           =   1125
            End
            Begin VB.Label lblIdEstablecimientoOrigen 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Estab. referencia"
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
               TabIndex        =   79
               Top             =   705
               Width           =   1380
            End
            Begin VB.Label lblIdTipoReferenciaOrigen 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo referencia"
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
               Height          =   210
               Left            =   120
               TabIndex        =   78
               Top             =   360
               Width           =   1230
            End
         End
         Begin VB.Frame fraDatosCita 
            Caption         =   "Cita"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3525
            Left            =   120
            TabIndex        =   49
            Top             =   330
            Width           =   5775
            Begin VB.TextBox txtFechaIngreso 
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
               Left            =   1650
               TabIndex        =   63
               Top             =   2040
               Width           =   1095
            End
            Begin VB.TextBox txtMedico 
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
               Left            =   1650
               TabIndex        =   62
               Top             =   960
               Width           =   3915
            End
            Begin VB.TextBox txtEdadEnDias 
               BackColor       =   &H00FFFFFF&
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
               Left            =   1650
               TabIndex        =   61
               Top             =   2400
               Width           =   645
            End
            Begin VB.TextBox txtHoraInicio 
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
               Left            =   3510
               TabIndex        =   60
               Top             =   2040
               Width           =   630
            End
            Begin VB.TextBox txtHoraFin 
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
               Left            =   4890
               TabIndex        =   59
               Top             =   2040
               Width           =   690
            End
            Begin VB.ComboBox cmbIdTipoServicio 
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
               Left            =   1650
               TabIndex        =   58
               Top             =   240
               Width           =   3930
            End
            Begin VB.ComboBox cmbIdViasAdmision 
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
               Left            =   1650
               TabIndex        =   57
               Top             =   600
               Width           =   3930
            End
            Begin VB.ComboBox cmbIdEspecialidadMedico 
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
               Left            =   1650
               TabIndex        =   56
               Top             =   1320
               Width           =   3930
            End
            Begin VB.ComboBox cmbIdServicio 
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
               Left            =   1650
               TabIndex        =   55
               Top             =   1680
               Width           =   3930
            End
            Begin VB.ComboBox cmbIdTipoEdad 
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
               Left            =   2370
               TabIndex        =   54
               Top             =   2400
               Width           =   1800
            End
            Begin VB.CheckBox chkDuplicadoCarne 
               Alignment       =   1  'Right Justify
               Caption         =   "Duplicado carnét"
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
               Left            =   3870
               TabIndex        =   53
               Top             =   3150
               Width           =   1695
            End
            Begin VB.CheckBox chkNuevoFolder 
               Alignment       =   1  'Right Justify
               Caption         =   "Nuevo folder"
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
               Left            =   2040
               TabIndex        =   52
               Top             =   3150
               Width           =   1425
            End
            Begin VB.CheckBox chkNuevoCarne 
               Alignment       =   1  'Right Justify
               Caption         =   "Nuevo carnét  "
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
               TabIndex        =   51
               Top             =   3180
               Width           =   1755
            End
            Begin VB.ComboBox cmbIdTipoConsulta 
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
               Left            =   1650
               TabIndex        =   50
               Top             =   2760
               Width           =   3960
            End
            Begin VB.Label lblFecha 
               BackStyle       =   0  'Transparent
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
               Height          =   315
               Left            =   120
               TabIndex        =   73
               Top             =   2070
               Width           =   1005
            End
            Begin VB.Label lblHoraInicio 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hora Ini"
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
               Left            =   2790
               TabIndex        =   72
               Top             =   2040
               Width           =   630
            End
            Begin VB.Label lblHoraFin 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hora Fin"
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
               Left            =   4200
               TabIndex        =   71
               Top             =   2055
               Width           =   660
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
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
               Left            =   90
               TabIndex        =   70
               Top             =   1710
               Width           =   915
            End
            Begin VB.Label Label43 
               BackStyle       =   0  'Transparent
               Caption         =   "Especialidad"
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
               TabIndex        =   69
               Top             =   1350
               Width           =   1065
            End
            Begin VB.Label Label44 
               BackStyle       =   0  'Transparent
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
               Left            =   90
               TabIndex        =   68
               Top             =   990
               Width           =   1365
            End
            Begin VB.Label lblEdadEnDias 
               BackStyle       =   0  'Transparent
               Caption         =   "Edad en la Atenc"
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
               TabIndex        =   67
               Top             =   2430
               Width           =   1455
            End
            Begin VB.Label lblIdTipoServicio 
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo de servicio"
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
               TabIndex        =   66
               Top             =   285
               Width           =   1395
            End
            Begin VB.Label lblViaAdmision 
               BackStyle       =   0  'Transparent
               Caption         =   "Origen"
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
               TabIndex        =   65
               Top             =   630
               Width           =   1155
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo Consulta"
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
               TabIndex        =   64
               Top             =   2790
               Width           =   1185
            End
         End
         Begin VB.Frame Frame15 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3945
            Left            =   -74850
            TabIndex        =   33
            Top             =   480
            Width           =   11595
            Begin VB.CommandButton Command7 
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
               Height          =   315
               Left            =   1470
               TabIndex        =   41
               Top             =   3510
               Width           =   465
            End
            Begin VB.CommandButton Command6 
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
               Height          =   315
               Left            =   1470
               TabIndex        =   40
               Top             =   2862
               Width           =   465
            End
            Begin VB.CommandButton Command5 
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
               Height          =   315
               Left            =   1470
               TabIndex        =   39
               Top             =   2214
               Width           =   465
            End
            Begin VB.CommandButton Command4 
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
               Height          =   315
               Left            =   1470
               TabIndex        =   38
               Top             =   1566
               Width           =   465
            End
            Begin VB.CommandButton Command3 
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
               Height          =   315
               Left            =   1470
               TabIndex        =   37
               Top             =   918
               Width           =   465
            End
            Begin VB.CommandButton Command2 
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
               Height          =   315
               Left            =   1470
               TabIndex        =   36
               Top             =   270
               Width           =   465
            End
            Begin VB.TextBox Text4 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3525
               Left            =   2130
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   35
               Top             =   270
               Width           =   9015
            End
            Begin VB.CommandButton Command1 
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
               Left            =   11130
               Style           =   1  'Graphical
               TabIndex        =   34
               Top             =   300
               Width           =   405
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Farmacia"
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
               Height          =   210
               Left            =   180
               TabIndex        =   47
               Top             =   3540
               Width           =   690
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Patol.Clínica"
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
               Height          =   210
               Left            =   180
               TabIndex        =   46
               Top             =   2898
               Width           =   945
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Ecog.General"
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
               Height          =   210
               Left            =   180
               TabIndex        =   45
               Top             =   2256
               Width           =   1080
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Ecog.Obstét"
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
               Height          =   210
               Left            =   180
               TabIndex        =   44
               Top             =   1614
               Width           =   1035
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Tomografía"
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
               Height          =   210
               Left            =   180
               TabIndex        =   43
               Top             =   972
               Width           =   915
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Rayos X"
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
               Height          =   210
               Left            =   180
               TabIndex        =   42
               Top             =   330
               Width           =   630
            End
         End
         Begin VB.Frame Frame14 
            Caption         =   "Tratamiento"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3945
            Left            =   -74820
            TabIndex        =   31
            Top             =   450
            Width           =   5655
            Begin VB.TextBox Text3 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3465
               Left            =   150
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   32
               Top             =   300
               Width           =   5295
            End
         End
         Begin VB.Frame Frame13 
            Caption         =   "Observaciones"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3945
            Left            =   -68940
            TabIndex        =   29
            Top             =   450
            Width           =   5655
            Begin VB.TextBox Text2 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3465
               Left            =   150
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   30
               Top             =   300
               Width           =   5295
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "Diagnóstico del Médico"
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
            Left            =   -74880
            TabIndex        =   27
            Top             =   3450
            Width           =   11745
            Begin VB.TextBox Text1 
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
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   28
               Top             =   240
               Width           =   11505
            End
         End
         Begin SISGalenPlus.ucDiagnosticoDetalle ucDiagnosticoDetalle2 
            Height          =   2985
            Left            =   -74880
            TabIndex        =   48
            Top             =   420
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   5265
         End
         Begin SISGalenPlus.UcEpisodioClinico UcEpisodioClinico1 
            Height          =   585
            Left            =   8160
            TabIndex        =   106
            Top             =   5610
            Visible         =   0   'False
            Width           =   3435
            _ExtentX        =   6059
            _ExtentY        =   1032
         End
         Begin SISGalenPlus.ucCitasLista ucCitasLista11 
            Height          =   5670
            Left            =   -74925
            TabIndex        =   127
            Top             =   345
            Width           =   11625
            _ExtentX        =   11086
            _ExtentY        =   8599
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Elegir el CUPO y pulsar ENTER para imprimir la CITA"
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
            Left            =   -74880
            TabIndex        =   128
            Top             =   6135
            Width           =   4635
         End
      End
   End
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
      Height          =   1965
      Left            =   20
      TabIndex        =   16
      Top             =   -30
      Width           =   7875
      Begin VB.CommandButton cmdSinApellidoMaterno 
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
         Left            =   2295
         Picture         =   "AdmisionDetalle.frx":184E
         Style           =   1  'Graphical
         TabIndex        =   120
         Top             =   990
         Width           =   300
      End
      Begin VB.CommandButton cmdSinApellidoPaterno 
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
         Left            =   1050
         Picture         =   "AdmisionDetalle.frx":1DD8
         Style           =   1  'Graphical
         TabIndex        =   119
         Top             =   975
         Width           =   300
      End
      Begin VB.CheckBox chkMuestraHistorial 
         Alignment       =   1  'Right Justify
         Caption         =   "Muestra HISTORIAL al buscar"
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
         Left            =   3915
         TabIndex        =   113
         Top             =   1545
         Width           =   2505
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   6510
         Picture         =   "AdmisionDetalle.frx":2362
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   1170
         Width           =   1305
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
         Left            =   2550
         MaxLength       =   20
         TabIndex        =   94
         Top             =   450
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Frame fraPacienteNuevo 
         Height          =   450
         Left            =   4050
         TabIndex        =   22
         Top             =   285
         Width           =   3735
         Begin VB.CheckBox chkBuscarEnSIS 
            Alignment       =   1  'Right Justify
            Caption         =   "Buscar en SIS"
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
            Left            =   30
            TabIndex        =   99
            Top             =   135
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.CheckBox chkPacienteNuevo 
            Alignment       =   1  'Right Justify
            Caption         =   "Paciente &nuevo"
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
            Left            =   2160
            TabIndex        =   23
            Top             =   135
            Width           =   1455
         End
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
         Left            =   1470
         MaxLength       =   9
         TabIndex        =   8
         Top             =   450
         Width           =   1065
      End
      Begin VB.CommandButton btnBuscarPaciente 
         Height          =   315
         Left            =   6510
         Picture         =   "AdmisionDetalle.frx":298B
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1500
         Width           =   1305
      End
      Begin VB.TextBox txtNroDNIBusqueda 
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
         TabIndex        =   7
         Top             =   450
         Width           =   1365
      End
      Begin VB.TextBox txtSegundoNombreBusqueda 
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
         Left            =   1410
         MaxLength       =   40
         TabIndex        =   3
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtApellidoMaternoBusqueda 
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
         Left            =   1410
         MaxLength       =   40
         TabIndex        =   1
         Top             =   990
         Width           =   870
      End
      Begin VB.TextBox txtApellidoPaternoBusqueda 
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
         MaxLength       =   40
         TabIndex        =   0
         Top             =   990
         Width           =   915
      End
      Begin VB.TextBox txtPrimerNombreBusqueda 
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
         MaxLength       =   40
         TabIndex        =   2
         Top             =   1560
         Width           =   1260
      End
      Begin VB.Label Label7 
         Caption         =   "1er Nombre          2do Nombre"
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
         Left            =   120
         TabIndex        =   109
         Top             =   1320
         Width           =   2505
      End
      Begin VB.Label lblFichaFamiliar 
         Caption         =   "Ficha Familiar"
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
         Left            =   2640
         TabIndex        =   95
         Top             =   240
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label6 
         Caption         =   "DNI                         N° Historia"
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
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   2445
      End
      Begin VB.Label Label50 
         Caption         =   "Apellido Paterno    ApellidoMaterno"
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
         Left            =   120
         TabIndex        =   17
         Top             =   780
         Width           =   3825
      End
   End
   Begin VB.Frame Frame4 
      Height          =   885
      Left            =   0
      TabIndex        =   15
      Top             =   8790
      Width           =   11895
      Begin VB.CommandButton btnImprimeFiliacion 
         Caption         =   "Filiación Arch.Clínico"
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
         Height          =   615
         Left            =   2670
         Picture         =   "AdmisionDetalle.frx":55D4
         Style           =   1  'Graphical
         TabIndex        =   108
         Top             =   180
         Width           =   1245
      End
      Begin VB.CommandButton btnImprimeFichaSIS 
         Caption         =   "FUA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9330
         Picture         =   "AdmisionDetalle.frx":5AAD
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   180
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.CommandButton btnBuscaHistoricos 
         Caption         =   "Históricos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10590
         Picture         =   "AdmisionDetalle.frx":5F86
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   180
         Width           =   1245
      End
      Begin VB.CommandButton btnImprimePreCta 
         Caption         =   "Cuenta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1380
         Picture         =   "AdmisionDetalle.frx":6510
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   180
         Width           =   1245
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar "
         DisabledPicture =   "AdmisionDetalle.frx":69E9
         DownPicture     =   "AdmisionDetalle.frx":6EAD
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6090
         Picture         =   "AdmisionDetalle.frx":7399
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   180
         Width           =   1245
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "AdmisionDetalle.frx":7885
         DownPicture     =   "AdmisionDetalle.frx":7CE5
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4770
         Picture         =   "AdmisionDetalle.frx":815A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   180
         Width           =   1245
      End
      Begin VB.CommandButton btnImprimir 
         Caption         =   "Hoja Filiación Consultorio"
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
         Height          =   615
         Left            =   120
         Picture         =   "AdmisionDetalle.frx":85CF
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   180
         Width           =   1245
      End
      Begin SISGalenPlus.ucMensajeParpadeando ucMensajeParpadeando1 
         Height          =   615
         Left            =   7350
         TabIndex        =   107
         Top             =   150
         Visible         =   0   'False
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   1085
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   1155
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
      Begin SISGalenPlus.ucFacturacionItems ucFacturacionProductos 
         Height          =   630
         Left            =   180
         TabIndex        =   20
         Top             =   210
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   1111
      End
   End
   Begin UltraGrid.SSUltraGrid grdAnteriores 
      Height          =   1860
      Left            =   7920
      TabIndex        =   21
      ToolTipText     =   "ROJO (no se registro la atención o Paciente faltó a CITA),   AZUL(atenciones mayores a HOY)"
      Top             =   30
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   3281
      _Version        =   131072
      GridFlags       =   17040388
      UpdateMode      =   2
      LayoutFlags     =   67108884
      RowConnectorColor=   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "grdAnteriores"
   End
   Begin VB.Image pi_imagen 
      BorderStyle     =   1  'Fixed Single
      Height          =   2295
      Left            =   0
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      ToolTipText     =   "Pulsar Doble Click para ampliar Imagen"
      Top             =   0
      Visible         =   0   'False
      Width           =   2745
   End
End
Attribute VB_Name = "AdmisionCEDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de CITAS
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'

'---Variable para guardar fua automatico

'HRA 10/12/2020 Cambio 47 Inicio
Dim GuardaFua As String
'Dim Codpr As New ReglasAdmision
'HRA 10/12/2020 Cambio 47 Fin


'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ReporteUtil As New sighEntidades.ReporteUtil
Dim mo_Teclado As New sighEntidades.Teclado
Dim mo_Formulario As New sighEntidades.Formulario
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim ml_idUsuario As Long
Dim mb_ExistenDatos As Boolean

'
Dim mo_ReglasSISgalenhos As New SIGHSis.ReglasSISgalenhos
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminServiciosGeograficos As New SIGHNegocios.ReglasServGeograf
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
Dim mo_SisConsumoWeb As New SIGHNegocios.SisConsumoWeb
Dim ml_TipoServicio As sghTipoServicio
Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_Reniec As New ReniecGalenhos
'
Dim mo_Apariencia As New sighEntidades.GridInfragistic
Dim ml_TipoVistaForm As sghTipoVistaFormAtenciones
Dim ml_EstadoCuenta As Long
'
Dim mo_cmbIdTipoServicio As New sighEntidades.ListaDespleglable
Dim mo_cmbIdViasAdmision As New sighEntidades.ListaDespleglable
Dim mo_cmbIdEspecialidadMedico As New sighEntidades.ListaDespleglable
Dim mo_cmbIdServicio As New sighEntidades.ListaDespleglable
Dim mo_cmbIdTipoReferenciaOrigen As New sighEntidades.ListaDespleglable
Dim mo_cmbIdTipoEdad As New sighEntidades.ListaDespleglable
Dim mo_cmbIdTipoConsulta As New sighEntidades.ListaDespleglable
'
Dim mo_cmbIdFormaPago As New sighEntidades.ListaDespleglable
Dim mo_cmbIdFuentesFinanciamiento As New sighEntidades.ListaDespleglable
'
Dim mo_Especialidad As New DOEspecialidades
Dim mo_paciente As New doPaciente
Dim mo_DoUbicacionPaciente As New doPaciente
Dim mo_DoAtencionDatosAdicionales As New DoAtencionDatosAdicionales
Dim mb_FormLoading As Boolean
Dim mo_FacturacionServicios As New Collection
Dim mo_FacturacionBienesInsumos As New Collection
Dim mo_FacturacionServiciosPorEliminar As New Collection
Dim mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String
'------------------------------------------------------------------------------------
'                               PACIENTE NUEVO -debb
'------------------------------------------------------------------------------------
Dim lcApP As String
Dim lcApM As String
Dim lcPnom As String
Dim lcSnombreReniec As String, ldFnacimientoReniec As Date, lnIdSexoReniec As Long
Dim lcDireccionReniec As String, mb_UsoWebReniec As Boolean
Dim lnIdDistritoSIS As Long, lnIdSexoSIS As Long, ldFechaNacimientoSIS As Date, lcSnombreSIS As String
Dim lnIdPlanSIS As Long, lcDniSIS As String, lnAfiliacionSIS1 As String, lnAfiliacionSIS2 As String, lnAfiliacionSIS5 As String
Dim lnAfiliacionSIS3 As String, lnAfiliacionSIS4 As Long, lcSIScodigo As String, lcTipoFormatoSIS As String
Dim lcCodigoEstablecimientoAdscripcionSIS As String, lbEncontroAfiliadoEnWebSIS As Boolean
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim oRsFormaPago As New ADODB.Recordset
Dim oRsFuentesFinanciamiento As New ADODB.Recordset
Dim oRsServiciosIntermedios As New Recordset
Dim rsServicio As New Recordset
Dim lnFormaPagoAnterior As Long
Dim lnIdFactServicios As Long
Dim lcFbajaok As String 'Actualiado 16102014

'------------------------------------------------------------------------------------
'                               VARIABLES CUENTAS DE ATENCION
'------------------------------------------------------------------------------------
Dim mo_CuentasAtencion As New DOCuentaAtencion
Dim ml_idCuentaAtencion As Long

'------------------------------------------------------------------------------------
'                               VARIABLE PARA LA ATENCION
'------------------------------------------------------------------------------------
Dim mo_Atenciones As New DOAtencion
Dim ml_idAtencion As Long
Dim mo_Diagnosticos As New Collection
Dim mo_Procedimientos As New Collection
Dim mo_Examenes As New Collection
'------------------------------------------------------------------------------------
'                               VARIABLE PARA LA FILIACION
'------------------------------------------------------------------------------------
Dim ml_IdPaciente As Long
Dim mo_Pacientes  As New doPaciente
Dim ms_Autogenerado As String
Dim ml_TipoNumeracion As sghTipoNumeracionDeNroHistoria
Dim mo_Historia As New DOHistoriaClinica
Dim mo_DoPacientesDatosAdd As New DoPacienteDatosAdd
'
'------------------------------------------------------------------------------------
'                               VARIABLE PARA RECETAS
'------------------------------------------------------------------------------------
Dim lnRecetaRayosX As Long, lnRecetaEcografiaO As Long, lnRecetaEcografiaG As Long
Dim lnRecetaTomografia As Long, lnRecetaAnatomiaP As Long, lnRecetaPatologiaC As Long
Dim lnRecetaBancoS As Long, lnRecetaFarmacia As Long

'------------------------------------------------------------------------------------
'                               VARIABLE PARA LA CITA
'------------------------------------------------------------------------------------

Dim mda_FechaIngreso As Date
Dim ms_HoraInicio As String
Dim ms_HoraFin As String
Dim ms_NombrePaciente As String
Dim mo_Cita As New DOCita
Dim mo_DOFacturacionPaquetes As New DOFacturacionPaquetes
Dim mo_DOFacturacionPaquetesAnt As New DOFacturacionPaquetes

Dim lcNserieAnt As String, lcNboletaAnt As String
Dim ml_IdMedico As Long
Dim ms_NombreMedico  As String
Dim ml_IdCita As Long
Dim ml_IdEstadoCita As Long
Dim ml_IdPrestamo As Long
Dim ml_IdProgramacion As Long
Dim idFormaPagoProvisional As Long
Dim ms_NroCola As String
Dim lbUsuarioAutorizadoAregistrarCitasRepetidas As Boolean
Dim mo_lbCargaTablasUnaVez As Boolean
Dim mo_lbNuevoMovimiento As Boolean
Dim lbYaSeTransfirioHCdeUnServicioAotro As Boolean
Dim mo_DOAtencionesCE As New DOAtencionesCE
Const lcLinea As String = "----------------------------------------------------------------------------------------"
Const lcLineaChar As String = "¨"
Dim lcHistoriaYpaciente As String
Dim oDoSunasaPacientesHistoricos As New DoSunasaPacientesHistoricos
Dim oDoPacienteDatosAdd As New DoPacienteDatosAdd
Dim lc_AntecedentePersonal As String
Dim mb_NecesitaTriaje As Boolean
Dim ml_FechaReceta As Date
Dim lbBuscaDNIenReniec As Boolean
Dim lbPacienteDatosAdicionalesEsNuevo As Boolean
Dim ldFechaActualServidor As Date
Dim lbElConsultorioUsaModuloPerinatal As Boolean
Dim lbElConsultorioUsaModuloMaterno As Boolean
Dim lbElMedicoNOregistraFUA As String
Dim lbCargaUnaSolaVez As Boolean
Dim mo_lbEsCitaAdicional As Boolean, lbCargaUnaVezVEntana As String
Const lbCargaAlaVezCitaPacienteAtencionDA As Boolean = False
Const lcPagoCita As String = "Pagada"
Dim wxParametroBusqRapida As String
'mgaray201503
Dim bEsCajero As Boolean
Dim mc_FuaVersionFormato As String
Dim mo_lcLlegoAlMaximoCuposSIS As String       'debb-25/08/2016
Dim mi_nroHistoriaCitadoXmedico As Long
Dim lbEsUnEPSdesdeAgregarCE As Boolean
Dim lbTieneLicenciaParaMensajeAcelulares As Boolean
Dim lbImpresionCuenta As Boolean
Dim lbElConsultorioNoEnviaMensajeTextoCelular As Boolean
Dim wxParametro580 As String, lbUsuarioTrabajaCitasPorTelefono As Boolean, lbTieneLicenciaTerapias As Boolean
Dim ml_idFuenteFinanciamientoCitadoXmedico As Long, ml_idFormaPagoCitadoXmedico As Long, ml_txtMedicoRefXMedico As String
Dim ml_cmbIdViasAdmisionXmedico As Long, ml_cmbIdTipoReferenciaOrigenXmedico As Long, ml_txtReferenciaOXmedico As String
Dim ml_txtIdEstablecimientoOrigenXmedico As String, ml_cmbServicioReferenciaOXmedico As String, ml_txtDxReferenciaXmedico As Long
Dim ml_lcCodigoEstablecimientoAdscripcionSISxMedico As String
Dim lnDocumentoTipoSIS As Long
Dim lbElConsultorioNoCobraApagantes As Boolean

Property Let lcCodigoEstablecimientoAdscripcionSISxMedico(lValue As String)
    ml_lcCodigoEstablecimientoAdscripcionSISxMedico = lValue
End Property

Property Let cmbIdViasAdmisionXmedico(lValue As Long)
    ml_cmbIdViasAdmisionXmedico = lValue
End Property
Property Let cmbIdTipoReferenciaOrigenXmedico(lValue As Long)
    ml_cmbIdTipoReferenciaOrigenXmedico = lValue
End Property
Property Let txtReferenciaOXmedico(lValue As String)
    ml_txtReferenciaOXmedico = lValue
End Property
Property Let txtIdEstablecimientoOrigenXmedico(lValue As String)
    ml_txtIdEstablecimientoOrigenXmedico = lValue
End Property
Property Let cmbServicioReferenciaOXmedico(lValue As String)
    ml_cmbServicioReferenciaOXmedico = lValue
End Property
Property Let txtDxReferenciaXmedico(lValue As Long)
    ml_txtDxReferenciaXmedico = lValue
End Property
Property Let txtMedicoRefXMedico(lValue As String)
    ml_txtMedicoRefXMedico = lValue
End Property

Property Let TieneLicenciaParaMensajeAcelulares(lValue As Boolean)
    lbTieneLicenciaParaMensajeAcelulares = lValue
End Property

Property Let idFormaPagoCitadoXmedico(lValue As Long)
   ml_idFormaPagoCitadoXmedico = lValue
End Property
Property Let idFuenteFinanciamientoCitadoXmedico(lValue As Long)
   ml_idFuenteFinanciamientoCitadoXmedico = lValue
End Property
'franklin 2017
Property Let nroHistoriaCitadoXmedico(lValue As Long)
    mo_lbCargaTablasUnaVez = True
    mi_nroHistoriaCitadoXmedico = lValue
End Property
Property Let LlegoAlMaximoCuposSIS(lValue As String)      'debb-25/08/2016
   mo_lcLlegoAlMaximoCuposSIS = lValue
End Property



Property Let EsCitaAdicional(lValue As Boolean)
   mo_lbEsCitaAdicional = lValue
End Property


Property Let lbNuevoMovimiento(lValue As Boolean)
   mo_lbNuevoMovimiento = lValue
End Property
Property Let lbCargaTablasUnaVez(lValue As Boolean)
   mo_lbCargaTablasUnaVez = lValue
End Property


Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
   
End Property
Property Let NroCola(lValue As Long)
   ms_NroCola = lValue
  ' fraDatosCita.Caption = "Cita N° " & ms_NroCola
End Property
Property Let IdCita(lValue As Long)
   ml_IdCita = lValue
End Property
Property Get IdCita() As Long
   IdCita = ml_IdCita
End Property
Property Let IdEstadoCita(lValue As Long)
   ml_IdEstadoCita = lValue
End Property
Property Get IdEstadoCita() As Long
   IdEstadoCita = ml_IdEstadoCita
End Property
Property Let idMedico(lValue As Long)
   ml_IdMedico = lValue
End Property
Property Get idMedico() As Long
   idMedico = ml_IdMedico
End Property
Property Let NombreMedico(sValue As String)
   ms_NombreMedico = sValue
End Property
Property Get NombreMedico() As String
   NombreMedico = ms_NombreMedico
End Property
Property Let FechaIngreso(lValue As Date)
   mda_FechaIngreso = lValue
End Property
Property Get FechaIngreso() As Date
   FechaIngreso = mda_FechaIngreso
End Property
Property Let HoraInicio(lValue As String)
   ms_HoraInicio = lValue
End Property
Property Get HoraInicio() As String
   HoraInicio = ms_HoraInicio
End Property
Property Let HoraFin(lValue As String)
   ms_HoraFin = lValue
End Property
Property Get HoraFin() As String
   HoraFin = ms_HoraFin
End Property
Property Let NombrePaciente(lValue As String)
   ms_NombrePaciente = lValue
End Property
Property Get NombrePaciente() As String
   NombrePaciente = ms_NombrePaciente
End Property
Property Set Cita(lValue As DOCita)
   Set mo_Cita = lValue
End Property
Property Get Cita() As DOCita
   Set Cita = mo_Cita
End Property
Property Let IdPrestamo(lValue As Long)
   ml_IdPrestamo = lValue
End Property
Property Get IdPrestamo() As Long
   IdPrestamo = ml_IdPrestamo
End Property
Property Let IdProgramacion(lValue As Long)
   ml_IdProgramacion = lValue
End Property
Property Get IdProgramacion() As Long
   IdProgramacion = ml_IdProgramacion
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
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let idCuentaAtencion(lValue As Long)
   ml_idCuentaAtencion = lValue
End Property
Property Get idCuentaAtencion() As Long
   idCuentaAtencion = ml_idCuentaAtencion
End Property
Property Let idAtencion(lValue As Long)
   ml_idAtencion = lValue
End Property
Property Get idAtencion() As Long
   idAtencion = ml_idAtencion
End Property
Property Let idPaciente(lValue As Long)
   ml_IdPaciente = lValue
End Property
Property Get idPaciente() As Long
   idPaciente = ml_IdPaciente
End Property
Property Let Autogenerado(sValue As String)
   ms_Autogenerado = sValue
End Property
Property Get Autogenerado() As String
   Autogenerado = ms_Autogenerado
End Property
Property Let TipoServicio(sValue As sghTipoServicio)
   ml_TipoServicio = sValue
End Property
Property Get TipoServicio() As sghTipoServicio
   TipoServicio = ml_TipoServicio
End Property
Property Let TipoNumeracion(lValue As Long)
   ml_TipoNumeracion = lValue
End Property
Property Get TipoNumeracion() As Long
   TipoNumeracion = ml_TipoNumeracion
End Property
Property Let TipoVistaForm(lValue As sghTipoVistaFormAtenciones)
   ml_TipoVistaForm = lValue
End Property

Sub CargarComboBoxes()
Dim sSQL As String
Dim sMensaje As String

        
        mo_cmbIdViasAdmision.BoundColumn = "IdOrigenAtencion"
        mo_cmbIdViasAdmision.ListField = "DescripcionLarga"
        Set mo_cmbIdViasAdmision.RowSource = mo_AdminAdmision.TiposOrigenAtencionSeleccionarViasDeConsultoriosExt
        sMensaje = sMensaje + mo_AdminAdmision.MensajeError
        
        
       mo_cmbIdTipoReferenciaOrigen.BoundColumn = "IdTipoReferencia"
       mo_cmbIdTipoReferenciaOrigen.ListField = "DescripcionLarga"
       Set mo_cmbIdTipoReferenciaOrigen.RowSource = mo_AdminServiciosComunes.TiposReferenciaSeleccionarTodos
       sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError
       
       
       mo_cmbIdTipoEdad.BoundColumn = "IdTipoEdad"
       mo_cmbIdTipoEdad.ListField = "DescripcionLarga"
       Set mo_cmbIdTipoEdad.RowSource = mo_AdminServiciosComunes.TiposEdadSeleccionarTodos
       sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError
       
        mo_cmbIdTipoServicio.BoundColumn = "IdTipoServicio"
        mo_cmbIdTipoServicio.ListField = "DescripcionLarga"
        Set mo_cmbIdTipoServicio.RowSource = mo_AdminServiciosHosp.TiposServicioSeleccionarAsistenciales
        mo_cmbIdTipoServicio.BoundText = "1"
        mo_Formulario.HabilitarDeshabilitar cmbIdTipoServicio, False
        
        Me.ucPacientesDetalle1.ConfigurarComboBoxes
        
        '
        Set oRsFormaPago = mo_AdminServiciosComunes.TiposFinanciamientoSegunFiltro("esFuenteFinanciamiento=1")
        mo_cmbIdFormaPago.BoundColumn = "idTipoFinanciamiento"
        mo_cmbIdFormaPago.ListField = "Descripcion"
        Set mo_cmbIdFormaPago.RowSource = oRsFormaPago
        mo_Formulario.HabilitarDeshabilitar Me.cmbFormaPago, False
        '
        Set oRsFuentesFinanciamiento = mo_AdminServiciosComunes.FuentesFinanciamientoSegunFiltro("UtilizadoEn=1 or UtilizadoEn=3")
        mo_cmbIdFuentesFinanciamiento.BoundColumn = "idFuenteFinanciamiento"
        mo_cmbIdFuentesFinanciamiento.ListField = "Descripcion"
        Set mo_cmbIdFuentesFinanciamiento.RowSource = oRsFuentesFinanciamiento
        '
        mo_Formulario.HabilitarDeshabilitar txtEdadEnDias, False
        mo_Formulario.HabilitarDeshabilitar cmbIdTipoEdad, False
        '----------------------------------------------------------------------------------
       sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError
       If sMensaje <> "" Then
           MsgBox mo_AdminServiciosComunes.MensajeError, vbInformation, Me.Caption
       End If
       'mgaray201503
       bEsCajero = UsuarioActualEsCajero()
       
       Set cmbServicioReferenciaO.ListSource = mo_AdminServiciosComunes.SuSalud_upsSeleccionarTodos   'debb-21/06/2016
End Sub

Private Sub btnBuscarEstablecimiento_Click()
    If cmbIdTipoReferenciaOrigen.Text <> "" Then
       CompletarDatosDeEstablecimiento txtIdEstablecimientoOrigen, txtNombreOrigenReferencia, mo_cmbIdTipoReferenciaOrigen.BoundText
    End If
End Sub

Sub BuscarPacientesSISSegunFiltro()
    Dim lcSql As String, oRsBuscaPacientesSis As New Recordset
    If UCase(Trim(lnAfiliacionSIS2)) = "R" Then
       lnAfiliacionSIS2 = UCase(Trim(lnAfiliacionSIS2))
    End If
    If (lnAfiliacionSIS1 <> "" And lnAfiliacionSIS2 <> "" And lnAfiliacionSIS3 <> "") Or _
       (lnAfiliacionSIS2 = "R" And lnAfiliacionSIS3 <> "") Or _
       (txtNroDNIBusqueda.Text <> "") Or (txtApellidoPaternoBusqueda.Text <> "") Then
       lcSql = ""
       If (lnAfiliacionSIS1 <> "" And lnAfiliacionSIS2 <> "" And lnAfiliacionSIS3 <> "") Then
          lcSql = "  where afiliacionDisa='" & lnAfiliacionSIS1 & "' and AfiliacionTipoFormato='" & lnAfiliacionSIS2 & _
                  "' and AfiliacionNroFormato='" & lnAfiliacionSIS3 & "' order by paterno,materno,pnombre"
       ElseIf lnAfiliacionSIS2 = "R" And lnAfiliacionSIS3 <> "" Then
          lcSql = "  where AfiliacionTipoFormato='" & lnAfiliacionSIS2 & "' and AfiliacionNroFormato='" & lnAfiliacionSIS3 & "' order by paterno,materno,pnombre"
       ElseIf txtNroDNIBusqueda.Text <> "" Then
          lcSql = "   where DocumentoTipo=1 and DocumentoNumero='" & txtNroDNIBusqueda.Text & "'"
       ElseIf txtApellidoPaternoBusqueda.Text <> "" Then
          lcSql = "   where  paterno like '%" & Trim(txtApellidoPaternoBusqueda.Text) & "%'"
          If txtApellidoMaternoBusqueda.Text <> "" Then
             lcSql = lcSql & " and  materno like '%" & Trim(txtApellidoMaternoBusqueda.Text) & "%'"
          End If
          If txtPrimerNombreBusqueda.Text <> "" Then
             lcSql = lcSql & " and  pnombre like '%" & Trim(txtPrimerNombreBusqueda.Text) & "%'"
          End If
          If txtSegundoNombreBusqueda.Text <> "" Then
             lcSql = lcSql & " and  onombres like '%" & Trim(txtSegundoNombreBusqueda.Text) & "%'"
          End If
       End If
       If lcSql <> "" Then
           lbEncontroAfiliadoEnWebSIS = False
           If wxParametro322 = "S" Then
              If (lnAfiliacionSIS1 <> "" And lnAfiliacionSIS2 <> "" And lnAfiliacionSIS3 <> "") Or _
                                             (lnAfiliacionSIS2 = "R" And lnAfiliacionSIS3 <> "") Then
                  '**************************Busca en Pag WEB del SIS x Nro Afiliado*******************
                  If Trim(lnAfiliacionSIS1) = "080" And Trim(lnAfiliacionSIS2) = "3" Then
                        lnAfiliacionSIS3 = Right("000000000" & Trim(lnAfiliacionSIS3), 9)
                  Else
                        lnAfiliacionSIS3 = Right("00000000" & Trim(lnAfiliacionSIS3), 8)
                  End If
                  Set oRsBuscaPacientesSis = mo_SisConsumoWeb.WebServiceSISBuscarAfiliado("", Trim(lnAfiliacionSIS1), _
                                                     Trim(lnAfiliacionSIS2), lnAfiliacionSIS3, _
                                                      lnAfiliacionSIS5, lcTipoFormatoSIS, wxParametro323)
                                                      
                  If oRsBuscaPacientesSis.RecordCount > 0 Then
                       lbEncontroAfiliadoEnWebSIS = True
                  End If
              ElseIf txtNroDNIBusqueda.Text <> "" Then
                  '***************************Busca en Pag WEB del SIS x DNI***************************
                  Set oRsBuscaPacientesSis = mo_SisConsumoWeb.WebServiceSISBuscarAfiliado(txtNroDNIBusqueda.Text, "", _
                                                      "", "", "", "", wxParametro323)
                  If oRsBuscaPacientesSis.RecordCount > 0 Then
                       lbEncontroAfiliadoEnWebSIS = True
                  End If
              End If
           End If
           If lbEncontroAfiliadoEnWebSIS = False Then
              If wxParametro322 = "S" And (txtNroDNIBusqueda.Text <> "" Or lnAfiliacionSIS3 <> "") And wxParametro526 <> "S" Then
                 Set oRsBuscaPacientesSis = mo_ReglasSISgalenhos.SisFiltraPacientesAfiliados(lcSql, wxParametroJAMO)
              ElseIf txtApellidoPaternoBusqueda.Text <> "" Or txtApellidoMaternoBusqueda.Text <> "" Then
                 Set oRsBuscaPacientesSis = mo_ReglasSISgalenhos.SisFiltraPacientesAfiliados(lcSql, wxParametroJAMO)
              End If
           End If
           If oRsBuscaPacientesSis.State = 0 Then
              Set oRsBuscaPacientesSis = Nothing
              Exit Sub
           End If
           Set grdPacientesEncontrados.DataSource = oRsBuscaPacientesSis.Clone
           
           'grdPacientesEncontrados.Caption = IIf(lbEncontroAfiliadoEnWebSIS = True, "Ubicado en la WEB SIS", wxParametro312)
           If lbEncontroAfiliadoEnWebSIS = False Then
              If wxParametro322 = "S" And (txtNroDNIBusqueda.Text <> "" Or lnAfiliacionSIS3 <> "") Then
                 grdPacientesEncontrados.Caption = wxParametro312 & "  (Verificar en el AREA DEL SIS SI ESTA AFILIADO, porque se buscó en la WEB SIS y no se encontró al Paciente)"
                 grdPacientesEncontrados.CaptionAppearance.ForeColor = vbRed
              Else
                 grdPacientesEncontrados.Caption = wxParametro312
                 grdPacientesEncontrados.CaptionAppearance.ForeColor = vbBlack
              End If
           Else
              grdPacientesEncontrados.Caption = "Ubicado en WEB SIS"
           End If
           
           Me.grdPacientesEncontrados.Visible = True
           grdPacientesEncontrados.Bands(0).Columns("cAfiliacion").Width = 1700
           grdPacientesEncontrados.Bands(0).Columns("EstadoSis").Width = 500
           grdPacientesEncontrados.Bands(0).Columns("fBAjaOk").Format = sighEntidades.DevuelveFechaSoloFormato_DMY
           If oRsBuscaPacientesSis.RecordCount = 1 Then
              grdPacientesEncontrados.SetFocus
           End If
           oRsBuscaPacientesSis.Close
       End If
    End If
    Set oRsBuscaPacientesSis = Nothing
End Sub

Private Sub btnBuscarPaciente_Click()
    If Me.chkBuscarEnSIS.Value = 1 Then
       Me.UcSISafiliacion1.TipoFormatoSISvisible False
       Me.UcSISafiliacion1.DevuelveValoresDeFiliacion lnAfiliacionSIS1, lnAfiliacionSIS2, lnAfiliacionSIS3, lcTipoFormatoSIS, lnAfiliacionSIS5
       BuscarPacientesSISSegunFiltro
       Exit Sub
    End If
   
    Dim RsHistorias As New Recordset
    Dim oDOPaciente As New doPaciente
    Dim oConexion As New Connection
    
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion
    
    
    
    lcApP = ""
    lcApM = ""
    lcPnom = ""
    lcDniSIS = ""
    lcSnombreSIS = ""
    lcSnombreReniec = "": ldFnacimientoReniec = 0: lnIdSexoReniec = 0: lcDireccionReniec = "": mb_UsoWebReniec = False
    If mo_Teclado.TextoEsSoloNumeros(Me.txtNroHistoriaBusqueda.Text) Then
'<(Inicio) Modificado Por: WABG el 16/10/2020-12:32:21 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
'      oDOPaciente.NroHistoriaClinica = Val(HCigualDNI_AgregaNUEVEaLaHistoria(Me.txtNroHistoriaBusqueda.Text))
       oDOPaciente.NroHistoriaClinica = Val(Me.txtNroHistoriaBusqueda.Text)
'</(Fin) Modificado Por: Project Administrator el 16/10/2020-12:32:21 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
    End If
    oDOPaciente.ApellidoPaterno = Me.txtApellidoPaternoBusqueda
    oDOPaciente.ApellidoMaterno = Me.txtApellidoMaternoBusqueda
    oDOPaciente.PrimerNombre = Me.txtPrimerNombreBusqueda
    oDOPaciente.SegundoNombre = Me.txtSegundoNombreBusqueda
'    If lnAfiliacionSIS3 = "" Then
       oDOPaciente.IdDocIdentidad = 1
'    Else
'      oDOPaciente.IdDocIdentidad = lnDocumentoTipoSIS
'    End If
    
    If lnDocumentoTipoSIS = 2 Then          'carnet extranjeria
        oDOPaciente.nrodocumento = ""
        oDOPaciente.IdDocIdentidad = 0
        Me.txtNroDNIBusqueda.Text = ""
    Else
         oDOPaciente.nrodocumento = Me.txtNroDNIBusqueda
    End If
    'oDOPaciente.nrodocumento = Me.txtNroDNIBusqueda
    
    If (oDOPaciente.ApellidoPaterno + oDOPaciente.ApellidoMaterno + _
    oDOPaciente.PrimerNombre + oDOPaciente.SegundoNombre = "") And _
    (Val(Me.txtNroHistoriaBusqueda.Text) = 0) And _
    (oDOPaciente.nrodocumento = "") And Me.txtFichaFamiliar.Text = "" Then
        MsgBox "Ingrese alguno de los valores de búsqueda", vbInformation, Me.Caption
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    If mo_Teclado.TextoEsSoloNumeros(Me.txtNroHistoriaBusqueda.Text) Then
       Set RsHistorias = mo_AdminAdmision.PacientesFiltraPorHistoriaClinicaDefinitiva(oDOPaciente, oConexion)
    ElseIf Me.txtFichaFamiliar.Text <> "" Then
       Set RsHistorias = mo_AdminAdmision.PacientesSeleccionarPorFichaFamiliar(Me.txtFichaFamiliar.Text)
    ElseIf Val(Me.txtNroDNIBusqueda.Text) > 0 Then
       Set RsHistorias = mo_AdminAdmision.PacientesFiltraPorNroDocumentoYtipo(oDOPaciente.nrodocumento, oDOPaciente.IdDocIdentidad, oConexion)
    Else
       If chkMuestraHistorial.Value = 1 Then
          Set RsHistorias = mo_AdminAdmision.PacientesFiltrarTodosSoloHistoriasDefinitivas(oDOPaciente, wxSinApellido, oConexion)
       Else
          Set RsHistorias = mo_AdminAdmision.PacientesFiltrarTodosSoloHistoriasDefinit_rap(oDOPaciente, wxSinApellido, oConexion)
       End If
    End If
    Screen.MousePointer = vbDefault
    
    grdPacientesEncontrados.Caption = "Lista de Pacientes encontrados en el ESTABLECIMIENTO"
    Set grdPacientesEncontrados.DataSource = RsHistorias
        
    'Si hay una sola coincidencia, además se buscó por DNI o HISTORIA
    If RsHistorias.RecordCount = 1 And (txtNroDNIBusqueda.Text <> "" Or txtNroHistoriaBusqueda.Text <> "") Then
        If mo_AdminAdmision.BuscaSiEstaHospitalizado(RsHistorias!idPaciente, oConexion, sghConsultaExterna) = False Or wxParametro539 = "S" Then    'debb-05/12/2015
            Me.grdPacientesEncontrados.Visible = False
            lcHistoriaYpaciente = "(" & RsHistorias!NroHistoriaClinica & ") " & Trim(RsHistorias!ApellidoPaterno) & _
                                 " " & Trim(RsHistorias!ApellidoMaterno) & " " & RsHistorias!PrimerNombre
            RsHistorias.MoveFirst
            chkPacienteNuevo.Value = 0
            Me.ucPacientesDetalle1.LimpiarDatosDePaciente wxParametro211, ldFechaActualServidor
            
            Me.ucPacientesDetalle1.idPaciente = RsHistorias!idPaciente
            Me.ucPacientesDetalle1.SegundoNombrePacienteSIS = ""
            Me.ucPacientesDetalle1.CargarDatosDePacienteALosControles oConexion, wxParametro242, wxParametro287
            
            Me.ucPacientesDetalle1.NroHistoriaClinica = RsHistorias!NroHistoriaClinica
            Me.ucPacientesDetalle1.TipoNumeracion = RsHistorias!idTipoNumeracion
            Me.idPaciente = RsHistorias!idPaciente
            Me.tabAdmision.Tab = 0
            MuestraCitasAnteriores oConexion, False
        End If
    ElseIf RsHistorias.RecordCount > 0 Then
        Me.grdPacientesEncontrados.Visible = True
        RsHistorias.MoveFirst
        If RsHistorias.RecordCount = 1 Then
           On Error Resume Next
           grdPacientesEncontrados.SetFocus
        End If
    ElseIf RsHistorias.RecordCount = 0 Then
        MsgBox "No se encontró datos" & Chr(13) & Chr(13) & "en la Base de Datos del Establecimiento", vbInformation, Me.Caption
        lcApP = txtApellidoPaternoBusqueda
        lcApM = txtApellidoMaternoBusqueda
        lcPnom = txtPrimerNombreBusqueda

        LimpiarFormulario
        Me.grdPacientesEncontrados.Visible = False
        txtNroHistoriaBusqueda.Text = ""
        txtApellidoMaternoBusqueda = ""
        txtPrimerNombreBusqueda = ""
        txtSegundoNombreBusqueda = ""
        txtApellidoPaternoBusqueda = ""
        txtNroDNIBusqueda = ""
    End If
    
    oConexion.Close
    Set oConexion = Nothing
    
    Set RsHistorias = Nothing
    Set oDOPaciente = Nothing
    
    Screen.MousePointer = vbDefault
ErrBP:
End Sub





Sub CargaCPTrealizadosEnElServicio()
    Set oRsServiciosIntermedios = mo_AdminAdmision.BuscaAtencionesCptCEparaFormatoHIS(Val(txtNroCuenta.Text), sghPuntosCargaBasicos.sghPtoCargaServicioHospitalizacion)
End Sub


Private Sub btnDxReferencia_Click()
    BusquedaDx ""
End Sub

Private Sub btnImprimeFichaSIS_Click()
    If mi_Opcion <> sghAgregar Then
       CargaDatosAlObjetosDeDatos
    End If
    If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE Then
         If ValidarDatosObligatorios = False Or ValidarReglas = False Then
               Exit Sub
         End If
    ElseIf mi_Opcion = sghModificar Then
         If ValidarDatosObligatorios = False Or ValidarReglas = False Or lbElMedicoNOregistraFUA = "N" Then
               Exit Sub
         End If
    End If
    Dim ml_FuaTipoAnexo2015 As Integer
    
    
    Dim oFua As New SIGHSis.clFUA
    oFua.CodigoPrestacion = Me.ucSISfuaCodPrestacion1.CodigoPrestacion
    oFua.idCuentaAtencion = Val(txtNroCuenta.Text) 'ml_idCuentaAtencion
    oFua.lcNombrePc = mo_lcNombrePc
    oFua.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
    oFua.idUsuario = ml_idUsuario
    oFua.Opcion = mi_Opcion
    oFua.IdServicio = CLng(Val(mo_cmbIdServicio.BoundText))
    
    'HRA 10/12/2020 Cambio 47 Inicio
    If GuardaFua = "F" Then
    oFua.MostrarFormularioF
    Else
    oFua.MostrarFormulario
    End If
    'HRA 10/12/2020 Cambio 47 Fin
  
     
    
    Set oFua = Nothing
End Sub

Private Sub btnImprimeFiliacion_Click()
    Dim oImprime As New RptHistoriaClinicaCE
    Dim oEdad As Edad
    If mi_Opcion = sghAgregar Then
       oEdad = sighEntidades.CalcularEdad(mo_Pacientes.FechaNacimiento, mo_Historia.fechacreacion)
    Else
       oEdad = sighEntidades.CalcularEdad(mo_Pacientes.FechaNacimiento, Me.ucPacientesDetalle1.DevuelveFechaCreacionHistoria)
    End If
    oImprime.ImprimeEnFormatoDeFiliacionParaHistoriaClinica mo_Atenciones.idPaciente, oEdad.Edad, oEdad.TipoEdad, Me.hwnd
    Set oImprime = Nothing
    Me.Visible = False
End Sub

Private Sub btnImprimePreCta_Click()
   If txtNroCuenta.Text <> "" Then
      lbImpresionCuenta = True
      ImprimePreCuenta
   End If
End Sub


Private Sub btnImprimir_Click()
    Dim oRptHistoriaConsultaExterna As New RptHistoriaClinicaCE
    If Me.idAtencion = 0 Then
        MsgBox "De agregar la atención para poder imprimir", vbInformation, Me.Caption
        Exit Sub
    End If
    oRptHistoriaConsultaExterna.idAtencion = Me.idAtencion
    oRptHistoriaConsultaExterna.idCuentaAtencion = Val(txtNroCuenta.Text)
    oRptHistoriaConsultaExterna.IdOrden = Val(txtNroOrdenPago.Text)
    oRptHistoriaConsultaExterna.CrearReporteHistoriaClinicaDeLaAtencionCE Me.hwnd
    Set oRptHistoriaConsultaExterna = Nothing
    
End Sub


Private Sub btnLimpiar_Click()
    txtNroDNIBusqueda.Text = ""
    txtNroHistoriaBusqueda.Text = ""
    txtApellidoPaternoBusqueda.Text = ""
    txtApellidoMaternoBusqueda.Text = ""
    txtPrimerNombreBusqueda.Text = ""
    txtSegundoNombreBusqueda.Text = ""
    txtFichaFamiliar.Text = ""
    UcSISafiliacion1.Limpiar
    'Actualizado Yamill Palomino 26102014
    ucMensajeParpadeando1.Visible = False
    Me.ucPacientesDetalle1.LimpiarDatosDePaciente wxParametro211, ldFechaActualServidor
    Set Me.grdAnteriores.DataSource = Nothing
    Me.grdPacientesEncontrados.Visible = False
    On Error Resume Next
    txtNroDNIBusqueda.SetFocus
End Sub





Private Sub chkBuscarEnSIS_Click()
    If chkBuscarEnSIS.Value = 1 Then
       fraBusqueda.Caption = "Solo se puede buscar en la WEB DEL SIS por DNI o por N°AFiliación"
       fraBusqueda.ForeColor = vbRed
       mo_Formulario.HabilitarDeshabilitar txtNroHistoriaBusqueda, False
       mo_Formulario.HabilitarDeshabilitar txtFichaFamiliar, False
       btnBuscarPaciente_Click
    Else
       fraBusqueda.Caption = "Búsqueda"
       fraBusqueda.ForeColor = vbBlack
       mo_Formulario.HabilitarDeshabilitar txtNroHistoriaBusqueda, True
       mo_Formulario.HabilitarDeshabilitar txtFichaFamiliar, True

    End If
End Sub



Private Sub chkPacienteNuevo_Click()
    Dim oConexion As New Connection
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    If chkPacienteNuevo.Value = 1 Then
        ucPacientesDetalle1.MarcoCheckPacienteNuevo = True
        LimpiarFormulario
        
        
        grdPacientesEncontrados.Visible = False
        grdAnteriores.Visible = False
        
        txtNroHistoriaBusqueda.Text = ""
        txtApellidoMaternoBusqueda = ""
        txtPrimerNombreBusqueda = ""
        txtSegundoNombreBusqueda = ""
        txtApellidoPaternoBusqueda = ""
        'txtNroDNIBusqueda = ""
        Me.tabAdmision.Tab = 0
        
'<(Inicio) Añadido Por: WABG el: 27/10/2020-08:00:01 p.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
        Me.ucPacientesDetalle1.HabilitarControlesDeTextoRENIEC
'</(Fin) Añadido Por: WABG el: 27/10/2020-08:00:01 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
        
        Me.ucPacientesDetalle1.ConfigurarValoresPorDefecto
        Me.ucPacientesDetalle1.Opcion = mi_Opcion
        'WCG_20060306
        chkNuevoCarne.Value = vbChecked
        chkNuevoFolder.Value = vbChecked
        chkNuevoCarne.Enabled = False
        chkNuevoFolder.Enabled = False
        chkDuplicadoCarne.Visible = False
        chkDuplicadoCarne.Value = vbUnchecked
        chkNuevoCarne.Visible = True
        chkNuevoFolder.Visible = True
        '
        If lbBuscaDNIenReniec = True And Len(txtNroDNIBusqueda.Text) = 8 Then
           mo_Reniec.ConsultarDNIenReniec txtNroDNIBusqueda.Text
           If mo_Reniec.ApellidoPaterno <> "" Then
                 lcApP = mo_Reniec.ApellidoPaterno
                 lcApM = mo_Reniec.ApellidoMaterno
                 lcPnom = mo_Reniec.PrimerNombre
                 lcSnombreReniec = mo_Reniec.SegundoNombre
                 ldFnacimientoReniec = mo_Reniec.FechaNacimiento
                 lnIdSexoReniec = mo_Reniec.idTipoSexo
                 lcDireccionReniec = mo_Reniec.DireccionDomicilio
                 mb_UsoWebReniec = True
           End If
        End If
        '
        If txtNroDNIBusqueda.Text = "" Then
           txtNroDNIBusqueda.Text = lcDniSIS
        End If
        '
        Me.ucPacientesDetalle1.CargaDatosBasicosPacienteNuevo UCase(lcApP), UCase(lcApM), UCase(lcPnom), wxParametro211, lcSnombreReniec, ldFnacimientoReniec, lnIdSexoReniec, lcDireccionReniec, mb_UsoWebReniec, txtNroDNIBusqueda.Text, lcSnombreSIS, lnIdDistritoSIS, lnIdSexoSIS, ldFechaNacimientoSIS
        txtNroDNIBusqueda.Text = ""
        '
        UcPacientesSunasa1.YaNoTieneSeguro
        '
        UcSISafiliacion1.InabilitaControles False
        If lnIdPlanSIS > 0 Then
             mo_cmbIdFuentesFinanciamiento.BoundText = lnIdPlanSIS
             cmbFuenteFinanciamiento_Click
        End If
        '
        Me.ucPacientesDetalle1.SetFocusEnDNI
    Else
        ucPacientesDetalle1.MarcoCheckPacienteNuevo = False
        '
        grdAnteriores.Visible = True
        idPaciente = 0
        MuestraCitasAnteriores oConexion, False
        'WCG_20060306
        chkNuevoCarne.Value = vbUnchecked
        chkNuevoFolder.Value = vbUnchecked
        chkNuevoCarne.Visible = False
        chkNuevoFolder.Visible = False
        chkDuplicadoCarne.Visible = True
        '
        UcSISafiliacion1.InabilitaControles True
    End If
    oConexion.Close
    Set oConexion = Nothing
    
    mo_Formulario.HabilitarDeshabilitar Me.txtNroHistoriaBusqueda, Not (chkPacienteNuevo.Value = 1)
    mo_Formulario.HabilitarDeshabilitar Me.txtApellidoMaternoBusqueda, Not (chkPacienteNuevo.Value = 1)
    mo_Formulario.HabilitarDeshabilitar Me.txtPrimerNombreBusqueda, Not (chkPacienteNuevo.Value = 1)
    mo_Formulario.HabilitarDeshabilitar Me.txtSegundoNombreBusqueda, Not (chkPacienteNuevo.Value = 1)
    mo_Formulario.HabilitarDeshabilitar Me.txtApellidoPaternoBusqueda, Not (chkPacienteNuevo.Value = 1)
    mo_Formulario.HabilitarDeshabilitar Me.txtNroDNIBusqueda, Not (chkPacienteNuevo.Value = 1)


End Sub

Private Sub chkPacienteNuevo_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub




Sub HaceVisibleBoleta()
    Dim oConexion As New Connection
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    lblBoleta.Visible = False
    txtNserie.Visible = False
    txtNboleta.Visible = False
    Me.UcPacientesSunasa1.HabilitaFrame True
    Me.UcPacientesSunasa1.YaNoTieneSeguro
    If Val(mo_cmbIdFormaPago.BoundText) > 0 Then
        If mo_AdminFacturacion.TiposFinanciamientoGeneraReciboPago(Val(mo_cmbIdFormaPago.BoundText), oConexion) = True Then
            lblBoleta.Visible = True
            txtNserie.Visible = True
            txtNboleta.Visible = True
            Me.UcPacientesSunasa1.HabilitaFrame False
        Else
            Me.UcPacientesSunasa1.idPaciente = ml_IdPaciente
            Me.UcPacientesSunasa1.CargarDatosDelUltimoSeguroDelPacienteALosControles oConexion
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
End Sub



Private Sub cmbFormaPago_Click()
    HaceVisibleBoleta
End Sub

Sub MuestraSiTieneEPSelFinancimientoElegido()
     On Error GoTo ErrSEPS
     ucEPS1.Visible = False
     oRsFuentesFinanciamiento.MoveFirst
     oRsFuentesFinanciamiento.Find "idFuenteFinanciamiento=" & mo_cmbIdFuentesFinanciamiento.BoundText
     If oRsFuentesFinanciamiento!tieneEPS = 1 Then
        ucEPS1.Visible = True
        If mi_Opcion = sghAgregar Then
           ucEPS1.Inicializar
        End If
     End If
ErrSEPS:
End Sub

Private Sub cmbFuenteFinanciamiento_Click()
        Set oRsFormaPago = mo_AdminFacturacion.TiposFinanciamientosTarifaSeleccionarPorPlan(Val(mo_cmbIdFuentesFinanciamiento.BoundText))
        Set mo_cmbIdFormaPago.RowSource = oRsFormaPago
        mo_cmbIdFormaPago.ListField = "Descripcion"
        mo_cmbIdFormaPago.BoundColumn = "idTipoFinanciamiento"
        '
        mo_Formulario.HabilitarDeshabilitar Me.cmbFormaPago, True
        If oRsFormaPago.RecordCount = 1 Then
           oRsFormaPago.MoveFirst
           mo_cmbIdFormaPago.BoundText = oRsFormaPago.Fields!idTipoFinanciamiento
           HaceVisibleBoleta
        ElseIf Val(mo_cmbIdFuentesFinanciamiento.BoundText) = 1 Then
           mo_cmbIdFormaPago.BoundText = wxParametro258
           HaceVisibleBoleta
        End If
        '
        UcPacientesSunasa1.idTipoFinanciamiento = Val(mo_cmbIdFormaPago.BoundText)
        '
        If UcSISafiliacion1.Visible = True Then
            HaceVisibleOnoBotonFUA
            If cmbFuenteFinanciamiento.Locked = False And Val(mo_cmbIdFuentesFinanciamiento.BoundText) = sghFuenteFinanciamiento.sghFFSIS And lnIdPlanSIS = 0 Then
               Dim lcDNI As String, lbPreguntar As Boolean
               If ucPacientesDetalle1.DevuelveFechaNacimiento <> sighEntidades.FECHA_VACIA_DMY Then 'Actualizado 16092014
                    If mo_ReglasSISgalenhos.PacienteBuscadoEnTablaGalenHosTieneAfiliacionSIS(ucPacientesDetalle1.DevuelveDNI, _
                                                 ucPacientesDetalle1.DevuelveApaterno, ucPacientesDetalle1.DevuelveAmaterno, _
                                                 ucPacientesDetalle1.DevuelvePnombre, ucPacientesDetalle1.DevuelveSnombre, _
                                                 ucPacientesDetalle1.DevuelveSexo, ucPacientesDetalle1.DevuelveFechaNacimiento, _
                                                 wxParametroJAMO, ldFechaActualServidor, lnAfiliacionSIS4, lcSIScodigo, True) = False Then
                            mo_cmbIdFuentesFinanciamiento.BoundText = ""
                            mo_cmbIdFormaPago.BoundText = ""
                    End If
               End If
               On Error Resume Next
               btnAceptar.SetFocus
            End If
        End If
        '
        MuestraSiTieneEPSelFinancimientoElegido
End Sub


Private Sub cmbFuenteFinanciamiento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       cmbFuenteFinanciamiento_Click
       btnAceptar.SetFocus
    End If
End Sub

Private Sub cmbIdTipoEdad_LostFocus()
Dim oDOTipoEdad As New DOTipoEdad

   If cmbIdTipoEdad.Text <> "" Then
     Set oDOTipoEdad = mo_AdminServiciosComunes.TiposEdadSeleccionarPorCodigo(Trim(Split(cmbIdTipoEdad.Text, " = ")(0)))
     If oDOTipoEdad.idTipoEdad <> 0 Then
         mo_cmbIdTipoEdad.BoundText = oDOTipoEdad.idTipoEdad
    End If
   End If
   Set oDOTipoEdad = Nothing
   mo_Formulario.MarcarComoVacio cmbIdTipoEdad
End Sub



Private Sub cmbServicioReferenciaO_KeyDown(KeyCode As Integer, Shift As Integer)
        mo_Teclado.RealizarNavegacion KeyCode, cmbServicioReferenciaO

End Sub

Private Sub cmdSinApellidoMaterno_Click()
    txtApellidoMaternoBusqueda.Text = wxSinApellido
End Sub

Private Sub cmdSinApellidoPaterno_Click()
    txtApellidoPaternoBusqueda.Text = wxSinApellido
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub

Private Sub grdAnteriores_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
       grdAnteriores.Bands(0).Columns("Fecha").Width = 800
       grdAnteriores.Bands(0).Columns("Consultorio").Width = 900
       grdAnteriores.Bands(0).Columns("Plan").Width = 900
       grdAnteriores.Bands(0).Columns("CS").Width = 300
       grdAnteriores.Bands(0).Columns("CE").Width = 300
       grdAnteriores.Bands(0).Columns("atendido").Width = 500
End Sub

Sub grdPacientesEncontradosSIS()
    On Error GoTo errGrdSis
    Dim oRecordset As New Recordset
    Dim rsPaciente As Recordset
    
    Dim lbValidaSiEsAfiliadoActualDelSIS As Boolean
    Dim lcSql As String
   
    Set oRecordset = grdPacientesEncontrados.DataSource
    lnIdDistritoSIS = 0: lnIdSexoSIS = 0: ldFechaNacimientoSIS = 0: lcSnombreSIS = "":  lcDniSIS = ""
    lnIdPlanSIS = sghFuenteFinanciamiento.sghFFSIS
    lcCodigoEstablecimientoAdscripcionSIS = ""
    lcFbajaok = ""
    If oRecordset.RecordCount > 0 Then
        If mo_ReglasSISgalenhos.Sis_ValidaSiEsAfiliadoActualDelSIS(oRecordset, CDate(Me.txtFechaIngreso.Text), True) = True Then
            lnAfiliacionSIS1 = oRecordset.Fields!cDisa
            lnAfiliacionSIS2 = oRecordset.Fields!cFormato
            lnAfiliacionSIS3 = oRecordset.Fields!cnumero
            lnAfiliacionSIS4 = oRecordset.Fields!idSiaSis
            lcSIScodigo = oRecordset.Fields!Codigo
            
            lnDocumentoTipoSIS = oRecordset.Fields!DocumentoTipo
            
            
            lcDniSIS = IIf(IsNull(oRecordset.Fields!DNI), "", oRecordset.Fields!DNI)
            lcCodigoEstablecimientoAdscripcionSIS = IIf(IsNull(oRecordset.Fields!CodigoEstablAdscripcion), "", oRecordset.Fields!CodigoEstablAdscripcion)
            
            lcFbajaok = IIf(IsNull(oRecordset.Fields!fBajaOK), "", oRecordset.Fields!fBajaOK) 'Actualizado 20102014
            
            If Not IsNull(oRecordset.Fields!sNombre) Then
               lcSnombreSIS = oRecordset.Fields!sNombre
            End If
            If Not IsNull(oRecordset.Fields!DistritoDomicilio) Then
               lnIdDistritoSIS = Val(oRecordset.Fields!DistritoDomicilio)
            End If
            If Not IsNull(oRecordset.Fields!Sexo) Then
               lnIdSexoSIS = IIf(oRecordset.Fields!Sexo = "0", 2, 1)
            End If
            If Not IsNull(oRecordset.Fields!FNacimiento) Then
               ldFechaNacimientoSIS = oRecordset.Fields!FNacimiento
            End If
            txtNroDNIBusqueda.Text = IIf(IsNull(oRecordset.Fields!DNI), "", oRecordset.Fields!DNI)
            Me.txtApellidoPaternoBusqueda.Text = oRecordset.Fields!apPaterno
            Me.txtApellidoPaternoBusqueda.Text = oRecordset.Fields!apPaterno
            Me.txtApellidoMaternoBusqueda.Text = oRecordset.Fields!apMaterno
            Me.txtPrimerNombreBusqueda.Text = oRecordset.Fields!Pnombre
            Me.txtSegundoNombreBusqueda.Text = IIf(IsNull(oRecordset.Fields!sNombre), "", oRecordset.Fields!sNombre)
            Me.chkBuscarEnSIS.Value = 0
            btnBuscarPaciente_Click
            If Not IsNull(oRecordset.Fields!sNombre) Then
               lcSnombreSIS = oRecordset.Fields!sNombre
            End If
            lcDniSIS = IIf(IsNull(oRecordset.Fields!DNI), "", oRecordset.Fields!DNI)
            Set rsPaciente = Me.grdPacientesEncontrados.DataSource
            If rsPaciente.RecordCount = 1 Then
                    grdPacientesEncontrados_DblClick
                    
            End If
            If Me.ucPacientesDetalle1.idPaciente > 0 Then
                If lnIdPlanSIS > 0 Then
                     mo_cmbIdFuentesFinanciamiento.BoundText = lnIdPlanSIS
                     cmbFuenteFinanciamiento_Click

'<(Inicio) Añadido Por: WABG el: 26/01/2021-11:55:26 a.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
                     Me.ucPacientesDetalle1.SetFocusEnHistoria
'</(Fin) Añadido Por: WABG el: 26/01/2021-11:55:26 a.m. en el Equipo: SISGALENPLUS-PC<CAMBIO-37>

                End If
            End If
        End If
    End If
    Set oRecordset = Nothing
    Set rsPaciente = Nothing
    
errGrdSis:
Exit Sub
Resume
End Sub

Private Sub grdAnteriores_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
        If Row.Cells("fecha").GetText() > ldFechaActualServidor Then
            Row.Appearance.ForeColor = vbBlue
        ElseIf Row.Cells("atendido").GetText() = "N" Then
            Row.Appearance.ForeColor = vbRed
        End If
End Sub

Private Sub grdPacientesEncontrados_DblClick()
    If Me.chkBuscarEnSIS.Value = 1 Then
       grdPacientesEncontradosSIS
       Exit Sub
    End If
    '
    Dim rsPaciente As Recordset
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion


    
    On Error Resume Next
    Set rsPaciente = Me.grdPacientesEncontrados.DataSource
    
    If mo_AdminAdmision.BuscaSiEstaHospitalizado(rsPaciente!idPaciente, oConexion, sghConsultaExterna) = True And wxParametro539 <> "S" Then  'debb-05/12/2015
       Exit Sub
    End If

    
    lcHistoriaYpaciente = "(" & rsPaciente!NroHistoriaClinica & ") " & Trim(rsPaciente!ApellidoPaterno) & " " & Trim(rsPaciente!ApellidoMaterno) & " " & rsPaciente!PrimerNombre
    Me.ucPacientesDetalle1.LimpiarDatosDePaciente wxParametro211, ldFechaActualServidor
    Me.ucPacientesDetalle1.TipoNumeracion = rsPaciente!idTipoNumeracion
    Me.ucPacientesDetalle1.NroHistoriaClinica = rsPaciente!NroHistoriaClnica
    Me.ucPacientesDetalle1.idPaciente = rsPaciente!idPaciente
    Me.idPaciente = rsPaciente!idPaciente
    Me.ucPacientesDetalle1.SegundoNombrePacienteSIS = lcSnombreSIS
    Me.ucPacientesDetalle1.CargarDatosDePacienteALosControles oConexion, wxParametro242, wxParametro287
    chkPacienteNuevo.Value = 0
    Me.tabAdmision.Tab = 0
    MuestraCitasAnteriores oConexion, False
    Me.grdPacientesEncontrados.Visible = False
    DoEvents
    Me.ucPacientesDetalle1.TabEnNroHistoria
    oConexion.Close
    Set oConexion = Nothing
    Set rsPaciente = Nothing
End Sub

Private Sub grdPacientesEncontrados_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    On Error Resume Next
    
    
    grdPacientesEncontrados.Bands(0).Columns("IdPaciente").Hidden = True
    grdPacientesEncontrados.Bands(0).Columns("IdTipoNumeracion").Hidden = True
    grdPacientesEncontrados.Bands(0).Columns("TipoNumeracion").Hidden = True
    
    grdPacientesEncontrados.Bands(0).Columns("NroHistoriaClinica").Header.Caption = "Nro Historia"
    grdPacientesEncontrados.Bands(0).Columns("NroHistoriaClinica").Width = 1000
    
    grdPacientesEncontrados.Bands(0).Columns("ApellidoPaterno").Header.Caption = "Ap. Paterno"
    grdPacientesEncontrados.Bands(0).Columns("ApellidoPaterno").Width = 1200
    
    grdPacientesEncontrados.Bands(0).Columns("ApellidoMaterno").Header.Caption = "Ap. Materno"
    grdPacientesEncontrados.Bands(0).Columns("ApellidoMaterno").Width = 1200
    
    grdPacientesEncontrados.Bands(0).Columns("PrimerNombre").Header.Caption = "1er Nombre"
    grdPacientesEncontrados.Bands(0).Columns("PrimerNombre").Width = 1200

    grdPacientesEncontrados.Bands(0).Columns("SegundoNombre").Header.Caption = "2do Nombre"
    grdPacientesEncontrados.Bands(0).Columns("SegundoNombre").Width = 1200

    grdPacientesEncontrados.Bands(0).Columns("FechaNacimiento").Header.Caption = "Fecha Nac."
    grdPacientesEncontrados.Bands(0).Columns("FechaNacimiento").Width = 1000

    grdPacientesEncontrados.Bands(0).Columns("TipoNumeracion").Header.Caption = "Tipo Numeración"
    grdPacientesEncontrados.Bands(0).Columns("TipoNumeracion").Width = 1500
    grdPacientesEncontrados.Bands(0).Columns("TipoNumeracion").CellAppearance.TextAlign = ssAlignRight

    grdPacientesEncontrados.Bands(0).Columns("TipoServicio").Header.Caption = "Ult.Tipo Serv."
    grdPacientesEncontrados.Bands(0).Columns("TipoServicio").Width = 1000

    grdPacientesEncontrados.Bands(0).Columns("FechaIngreso").Header.Caption = "Ult.Fec.Ing."
    grdPacientesEncontrados.Bands(0).Columns("FechaIngreso").Width = 1000

    grdPacientesEncontrados.Bands(0).Columns("FechaEgreso").Header.Caption = "Ult.Fec.Egr."
    grdPacientesEncontrados.Bands(0).Columns("FechaEgreso").Width = 1000

    grdPacientesEncontrados.Bands(0).Columns("ServicioIngreso").Header.Caption = "Ult.Serv.Ing."
    grdPacientesEncontrados.Bands(0).Columns("ServicioIngreso").Width = 2500
End Sub

Private Sub grdPacientesEncontrados_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    
    If KeyCode = vbKeyEscape Then
        grdPacientesEncontrados.Visible = False
    End If
    
End Sub

Private Sub grdPacientesEncontrados_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = vbKeyReturn Then
        grdPacientesEncontrados_DblClick
    End If
End Sub

Sub HabilitarFrameDestino(bValue As Boolean)
        mo_Formulario.HabilitarDeshabilitar btnBuscarEstablecimiento, bValue
End Sub



Sub CalculaLaHoraFinal()
Dim daHoraFin  As Date

    On Error Resume Next
    If Me.txtHoraInicio <> sighEntidades.HORA_VACIA_HM Then
        daHoraFin = DateAdd("n", mo_Especialidad.TiempoPromedioConsulta, CDate(Me.txtHoraInicio))
        Me.txtHoraFin = Format(daHoraFin, sighEntidades.DevuelveHoraSoloFormato_HM)
    Else
        Me.txtHoraFin = sighEntidades.HORA_VACIA_HM
    End If
    
End Sub





Private Sub TabIngreso_Click(PreviousTab As Integer)
   Select Case TabIngreso.Tab
   Case 1
        Me.UcPacientesSunasa1.DatosDeCabecera Me.ucPacientesDetalle1.DevuelvePaciente, Me.ucPacientesDetalle1.DevuelveSexo, _
            Me.ucPacientesDetalle1.DevuelveDocumento, Me.ucPacientesDetalle1.DevuelveNroDocumento, Me.ucPacientesDetalle1.DevuelvePaisDomicilio, _
            Me.ucPacientesDetalle1.DevuelveFechaNacimiento, Me.ucPacientesDetalle1.DevuelveUbigeoDomicilio
        '
        Me.ucCitasLista11.Visible = False
        If lbTieneLicenciaTerapias = True Then
            If mi_Opcion = sghAgregar And ml_IdPaciente > 0 And cmbFormaPago.Text <> "" Then
                Dim sMensaje As String
                sMensaje = ""
                CargaDatosAlObjetosDeDatos
                If Me.ValidarReglas And ValidaDatosObligatoriosReferencias(sMensaje) Then
                    Me.ucCitasLista11.Visible = True
                    Me.ucCitasLista11.lcCodigoEstablecimientoAdscripcionSISxMedico = lcCodigoEstablecimientoAdscripcionSIS
                    Me.ucCitasLista11.cmbIdViasAdmisionXmedico = Val(mo_cmbIdViasAdmision.BoundText)
                    Me.ucCitasLista11.cmbIdTipoReferenciaOrigenXmedico = Val(mo_cmbIdTipoReferenciaOrigen.BoundText)
                    Me.ucCitasLista11.txtReferenciaOXmedico = txtReferenciaO.Text
                    Me.ucCitasLista11.txtIdEstablecimientoOrigenXmedico = txtIdEstablecimientoOrigen.Tag
                    Me.ucCitasLista11.cmbServicioReferenciaOXmedico = PVcomboBoxDevuelveEleccion(cmbServicioReferenciaO)
                    Me.ucCitasLista11.txtDxReferenciaXmedico = Val(txtDxReferencia.Tag)
                    Me.ucCitasLista11.txtMedicoRefXMedico = Me.txtMedicoRef.Text
                    Me.ucCitasLista11.idFuenteFinanciamientoCitadoXmedico = Val(mo_cmbIdFuentesFinanciamiento.BoundText)
                    Me.ucCitasLista11.idFormaPagoCitadoXmedico = Val(mo_cmbIdFormaPago.BoundText)
                    Me.ucCitasLista11.IdMedicoAtencion = ml_IdMedico  ' mo_Atenciones.IdMedicoIngreso
                    Me.ucCitasLista11.nroHistoriaCitadoXmedico = Me.ucPacientesDetalle1.NroHistoriaClinica   'mo_paciente.NroHistoriaClinica
                    Me.ucCitasLista11.idPacienteCitadoXmedico = ml_IdPaciente 'mo_Atenciones.idPaciente
                    Me.ucCitasLista11.NOCargaDesdeCitas = True
                Else
                    MsgBox sMensaje, vbInformation, ""
                End If
            Else
                MsgBox "Solo funciona en AGREGAR, incluyendo PACIENTE con HISTORIA, eligiendo PRODUCTO/PLAN", vbInformation, ""
            End If
        End If
   End Select

End Sub



Private Sub tabAdmision_Click(PreviousTab As Integer)
    If tabAdmision.Tab = 1 Then
         On Error Resume Next
         TabIngreso.Tab = 0
         cmbFuenteFinanciamiento.SetFocus
    End If
End Sub



Private Sub tabAdmision_KeyDown(KeyCode As Integer, Shift As Integer)
AdministrarKeyPreview KeyCode
End Sub




Function EliminaAntecedentePersonal(lcMensajeEliminacion As String) As Boolean
    If MsgBox("Esta es una información médica registrada en al Base de Datos," & Chr(13) & _
              "si Ud. modifica la información, su USUARIO quedará grabado en el Sistema." & Chr(13) & Chr(13) & _
              "Esta seguro proseguir ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       EliminaAntecedentePersonal = True
       lc_AntecedentePersonal = "(" & lcMensajeEliminacion & ") "
    Else
       EliminaAntecedentePersonal = False
    End If
End Function
































Private Sub txtApellidoPaternoBusqueda_GotFocus()
lcApP = ""
End Sub



Private Sub txtDxReferencia_KeyDown(KeyCode As Integer, Shift As Integer)
     mo_Teclado.RealizarNavegacion KeyCode, txtDxReferencia
End Sub

Private Sub txtDxReferencia_LostFocus()
   If Len(txtDxReferencia.Text) > 0 And lblDxReferencia1.Text = "" Then
      BusquedaDx txtDxReferencia.Text
   End If
End Sub
'debb-21/06/2016
Sub BusquedaDx(lcCodigoDx As String)
    Dim oBusqueda As New SIGHNegocios.BuscaDiagnosticos
    Dim oDODiagnostico As DODiagnostico
    If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE Then
       oBusqueda.SoloMuestraDxGalenHos = False
    Else
       oBusqueda.SoloMuestraDxGalenHos = True
    End If
    oBusqueda.CodigoDx = lcCodigoDx
    oBusqueda.MostrarFormulario
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDODiagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorId(oBusqueda.idRegistroSeleccionado)
        If Not oDODiagnostico Is Nothing Then
            txtDxReferencia.Text = oDODiagnostico.CodigoCIE2004
            txtDxReferencia.Tag = oDODiagnostico.idDiagnostico
            lblDxReferencia1.Text = oDODiagnostico.descripcion
        Else
            txtDxReferencia.Text = ""
            txtDxReferencia.Tag = ""
            lblDxReferencia1.Text = ""
        End If
    Else
        txtDxReferencia.Text = ""
        txtDxReferencia.Tag = ""
        lblDxReferencia1.Text = ""
    End If
    Set oBusqueda = Nothing
End Sub


Private Sub txtFichaFamiliar_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFichaFamiliar
AdministrarKeyPreview KeyCode

End Sub

Private Sub txtFichaFamiliar_LostFocus()
   If txtFichaFamiliar.Text <> "" Then
      btnBuscarPaciente_Click
   End If

End Sub









Private Sub txtMedicoRef_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtMedicoRef
End Sub

Private Sub txtMedicoRef_LostFocus()
    BuscaMedicoRerencia ""
End Sub

Private Sub txtNboleta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txtNboleta_LostFocus
    End If
End Sub








Private Sub txtNroHistoriaBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNroHistoriaBusqueda
AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNroHistoriaBusqueda_KeyPress(KeyAscii As Integer)

    If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
    End If
    If KeyAscii = 13 And Len(txtNroHistoriaBusqueda.Text) > 0 Then
         btnBuscarPaciente_Click
    End If
End Sub


Private Sub cmbIdEspecialidadMedico_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdEspecialidadMedico
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdEspecialidadMedico_LostFocus()
   If cmbIdEspecialidadMedico.Text <> "" Then
        mo_cmbIdEspecialidadMedico.BoundText = Val(Split(cmbIdEspecialidadMedico.Text, " = ")(0))
        '
        Dim oRsTmp88 As New Recordset
        Dim lnDefaultTipoConsulta As Long
        lnDefaultTipoConsulta = 0
        Set oRsTmp88 = mo_AdminFacturacion.FacturacionSeleccionarTiposConsultaPorEspecialidad(Val(mo_cmbIdEspecialidadMedico.BoundText))
        If oRsTmp88.RecordCount > 0 Then
           oRsTmp88.MoveFirst
           Do While Not oRsTmp88.EOF
              If UCase(Left(oRsTmp88.Fields!nombre, 5)) <> "INTER" Then
                 lnDefaultTipoConsulta = oRsTmp88.Fields!idProducto
                 Exit Do
              End If
              oRsTmp88.MoveNext
           Loop
        End If
        oRsTmp88.Close
        Set oRsTmp88 = Nothing
        '
        mo_cmbIdTipoConsulta.BoundColumn = "IdProducto"
        mo_cmbIdTipoConsulta.ListField = "Descripcion"
        Set mo_cmbIdTipoConsulta.RowSource = mo_AdminFacturacion.FacturacionSeleccionarTiposConsultaPorEspecialidad(Val(mo_cmbIdEspecialidadMedico.BoundText))
        If lnDefaultTipoConsulta > 0 Then
           mo_cmbIdTipoConsulta.BoundText = lnDefaultTipoConsulta
        End If
   End If
   If Not cmbIdEspecialidadMedico.Locked Then mo_Formulario.MarcarComoVacio cmbIdEspecialidadMedico
End Sub

Private Sub cmbIdEspecialidadMedico_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub cmbIdServicio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdServicio
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdServicio_LostFocus()
   If cmbIdServicio.Text <> "" Then
       mo_cmbIdServicio.BoundText = Val(Split(cmbIdServicio.Text, " = ")(0))
   End If
   If Not cmbIdServicio.Locked Then mo_Formulario.MarcarComoVacio cmbIdServicio
End Sub

Private Sub cmbIdServicio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub






Private Sub cmbIdTipoReferenciaOrigen_Click()

    txtIdEstablecimientoOrigen.Tag = ""
    txtIdEstablecimientoOrigen = ""
    txtNombreOrigenReferencia = ""
End Sub

Private Sub cmbIdViasAdmision_Click()
Dim sCodigoOrigen As String
    If cmbIdViasAdmision.Text <> "" Then
       sCodigoOrigen = Trim(Split(cmbIdViasAdmision.Text, " = ")(0))
    Else
       sCodigoOrigen = " "
    End If
    If sCodigoOrigen <> "R" And sCodigoOrigen <> "C" Then
        mo_cmbIdTipoReferenciaOrigen.BoundText = ""
        Me.txtIdEstablecimientoOrigen = ""
        Me.txtIdEstablecimientoOrigen.Tag = ""
        cmbServicioReferenciaO.Text = "": lblDxReferencia1.Text = "": txtDxReferencia.Text = "": txtDxReferencia.Tag = ""  'debb-21/06/2016
        txtMedicoRef.Text = "": Me.cmbMedicoRef.Text = ""     'FRANKLIN 2017
    End If
    
    Select Case sCodigoOrigen
    Case "D", "A", "X", "E"
        HabilitarFrameOrigen False
        txtReferenciaO.Text = ""
    Case "R"
        HabilitarFrameOrigen True
        Me.fraDatosReferenciaOrigen = "Referencia origen "
        Me.lblIdTipoReferenciaOrigen = "Tipo Referencia"
        Me.lblIdEstablecimientoOrigen = "Estab. Referencia"
        mo_cmbIdTipoReferenciaOrigen.BoundText = "1"
    Case "C"
        HabilitarFrameOrigen True
        Me.fraDatosReferenciaOrigen = "Contrareferencia origen"
        Me.lblIdTipoReferenciaOrigen = "Tipo Contrarefer."
        Me.lblIdEstablecimientoOrigen = "Estab. Contrarefer."
        mo_cmbIdTipoReferenciaOrigen.BoundText = "1"
    End Select

End Sub
Sub HabilitarFrameOrigen(bValue As Boolean)
        mo_Formulario.HabilitarDeshabilitar fraDatosReferenciaOrigen, bValue
        mo_Formulario.HabilitarDeshabilitar fraDatosReferenciaOrigen, bValue
        mo_Formulario.HabilitarDeshabilitar lblIdTipoReferenciaOrigen, bValue
        mo_Formulario.HabilitarDeshabilitar cmbIdTipoReferenciaOrigen, bValue
        mo_Formulario.HabilitarDeshabilitar lblIdEstablecimientoOrigen, bValue
        mo_Formulario.HabilitarDeshabilitar btnBuscarEstablecimiento, bValue
        mo_Formulario.HabilitarDeshabilitar txtReferenciaO, bValue
        mo_Formulario.HabilitarDeshabilitar lblReferenciaO, bValue
        mo_Formulario.HabilitarDeshabilitar cmbServicioReferenciaO, bValue
        'debb-21/06/2016 (inicio)
        mo_Formulario.HabilitarDeshabilitar lblServicioReferencia, bValue
        mo_Formulario.HabilitarDeshabilitar lblDxReferencia1, False
        mo_Formulario.HabilitarDeshabilitar txtDxReferencia, bValue
        mo_Formulario.HabilitarDeshabilitar lblDxreferencia, bValue
        'debb-21/06/2016 (fin)
        'FRANKLIN 2017
        mo_Formulario.HabilitarDeshabilitar txtMedicoRef, bValue
        mo_Formulario.HabilitarDeshabilitar Me.cmbMedicoRef, bValue
        
End Sub
Private Sub cmbIdViasAdmision_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdViasAdmision
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdViasAdmision_LostFocus()
Dim oDOTipoOrigenAtencion As New DOTipoOrigenAtencion

   If cmbIdViasAdmision.Text <> "" Then
     Set oDOTipoOrigenAtencion = mo_AdminAdmision.TiposOrigenAtencionSeleccionarPorCodigo(Trim(Split(cmbIdViasAdmision.Text, " = ")(0)), ml_TipoServicio)
     If oDOTipoOrigenAtencion.IdOrigenAtencion <> 0 Then
         mo_cmbIdViasAdmision.BoundText = oDOTipoOrigenAtencion.IdOrigenAtencion
    End If
   End If
   mo_Formulario.MarcarComoVacio cmbIdViasAdmision

End Sub

Private Sub cmbIdViasAdmision_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub Form_Initialize()
    
    Set mo_cmbIdTipoServicio.MiComboBox = cmbIdTipoServicio
    Set mo_cmbIdViasAdmision.MiComboBox = cmbIdViasAdmision
    Set mo_cmbIdEspecialidadMedico.MiComboBox = cmbIdEspecialidadMedico
    Set mo_cmbIdServicio.MiComboBox = cmbIdServicio
    Set mo_cmbIdTipoReferenciaOrigen.MiComboBox = cmbIdTipoReferenciaOrigen
    Set mo_cmbIdTipoEdad.MiComboBox = cmbIdTipoEdad
    Set mo_cmbIdTipoConsulta.MiComboBox = cmbIdTipoConsulta
    Set mo_cmbIdFormaPago.MiComboBox = cmbFormaPago
    Set mo_cmbIdFuentesFinanciamiento.MiComboBox = cmbFuenteFinanciamiento
End Sub

Private Sub txtApellidoMaternoBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    'mo_Teclado.RealizarNavegacion KeyCode, txtApellidoMaternoBusqueda
    If KeyCode = 13 Then txtPrimerNombreBusqueda.SetFocus
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtApellidoMaternoBusqueda_LostFocus()
    txtApellidoMaternoBusqueda.Text = mo_Teclado.CapitalizarNombres(txtApellidoMaternoBusqueda.Text)
'   If Len(txtApellidoMaternoBusqueda.Text) > 0 Then
'      btnBuscarPaciente_Click
'   End If
End Sub

Private Sub txtApellidoMaternoBusqueda_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
   If KeyAscii = 13 Then
        If mo_AdminAdmision.RealizarBusquedaPacienteSiNo(wxParametroBusqRapida, txtApellidoPaternoBusqueda.Text, _
            txtApellidoMaternoBusqueda.Text) = True Then
            btnBuscarPaciente_Click
'        Else
'            If Trim(txtApellidoMaternoBusqueda.Text) = "" Then
'                MsgBox "Ingrese Apellido Materno a Buscar", vbInformation, "Mensaje"
'                txtApellidoMaternoBusqueda.SetFocus
'            Else
'                MsgBox "Ingrese Apellido Paterno a Buscar", vbInformation, "Mensaje"
'                txtApellidoPaternoBusqueda.SetFocus
'            End If
        End If
    End If
End Sub



Private Sub txtFechaIngreso_Change()
    
    On Error Resume Next
    Me.txtEdadEnDias = ""
    Dim oEdad As Edad
    oEdad = sighEntidades.CalcularEdad(CDate(Me.ucPacientesDetalle1.FechaNacimiento), CDate(txtFechaIngreso))
    Me.txtEdadEnDias = oEdad.Edad
    mo_cmbIdTipoEdad.BoundText = oEdad.TipoEdad

End Sub


Private Sub txtFechaIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFechaIngreso
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtFechaIngreso_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaFecha(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub



Private Sub txtHoraFin_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaHora(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtHoraInicio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtHoraInicio
    AdministrarKeyPreview KeyCode
End Sub




Private Sub txtHoraInicio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaHora(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub





Private Sub txtMedico_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtMedico
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNombreOrigenReferencia_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNombreOrigenReferencia
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtNombreOrigenReferencia_LostFocus()
    txtNombreOrigenReferencia.Text = mo_Teclado.CapitalizarNombres(txtNombreOrigenReferencia.Text)
End Sub

Private Sub txtNombreOrigenReferencia_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtNroDNIBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNroDNIBusqueda
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtNroDNIBusqueda_KeyPress(KeyAscii As Integer)
   If Len(txtNroDNIBusqueda.Text) = 8 And KeyAscii = 13 Then
      btnBuscarPaciente_Click
   ElseIf Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtApellidoPaternoBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
'   mo_Teclado.RealizarNavegacion KeyCode, txtApellidoPaternoBusqueda
    If KeyCode = 13 Then txtApellidoMaternoBusqueda.SetFocus
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtApellidoPaternoBusqueda_LostFocus()
    txtApellidoPaternoBusqueda.Text = mo_Teclado.CapitalizarNombres(txtApellidoPaternoBusqueda.Text)
'   If Len(txtApellidoPaternoBusqueda.Text) > 0 Then
'      btnBuscarPaciente_Click
'   End If
End Sub

Private Sub txtApellidoPaternoBusqueda_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
   If KeyAscii = 13 Then
        If mo_AdminAdmision.RealizarBusquedaPacienteSiNo(wxParametroBusqRapida, txtApellidoPaternoBusqueda.Text, _
                 txtApellidoMaternoBusqueda.Text) = True Then
           btnBuscarPaciente_Click
'        Else
'            If Trim(txtApellidoPaternoBusqueda.Text) = "" Then
'                MsgBox "Ingrese Apellido Paterno a Buscar", vbInformation, "Mensaje"
'                txtApellidoPaternoBusqueda.SetFocus
'            Else
'                MsgBox "Ingrese Apellido Materno a Buscar", vbInformation, "Mensaje"
'                txtApellidoMaternoBusqueda.SetFocus
'            End If
        End If
   End If
End Sub


Sub CargarDatosAlFormulario()

    Me.grdPacientesEncontrados.Left = 100 '210
    Me.grdPacientesEncontrados.Top = 1920 '1280
    mo_Formulario.HabilitarDeshabilitar Me.cmbIdTipoServicio, False
    mo_Formulario.HabilitarDeshabilitar Me.cmbFormaPago, False
    mo_Formulario.HabilitarDeshabilitar Me.cmbFuenteFinanciamiento, True
    
    HabilitarFrameOrigen False
    HabilitarFrameDestino False
    
    mo_Formulario.HabilitarDeshabilitar Me.txtFechaIngreso, False
    mo_Formulario.HabilitarDeshabilitar Me.txtHoraInicio, False
    mo_Formulario.HabilitarDeshabilitar Me.txtHoraFin, False
    
    mo_Formulario.HabilitarDeshabilitar Me.txtMedico, False
    mo_Formulario.HabilitarDeshabilitar Me.txtNroOrdenPago, False
    mo_Formulario.HabilitarDeshabilitar Me.txtNroCuenta, False
    mo_Formulario.HabilitarDeshabilitar txtNombreOrigenReferencia, False
    mo_Formulario.HabilitarDeshabilitar txtIdEstablecimientoOrigen, False
    

    Me.ucPacientesDetalle1.HacerVisibleCheckPacienteNoIdentificado False
    Me.ucPacientesDetalle1.NotaSobreUbicacion = "(*) Datos del día de la atención del paciente"
    
    
    
    Me.btnImprimir.Enabled = True
    Select Case mi_Opcion
     Case sghAgregar
        Me.btnImprimir.Enabled = False
        Me.ucPacientesDetalle1.TipoServicio = sghConsultaExterna
        ValoresPorDefecto
     Case sghModificar
         CargarDatosAlosControles
     Case sghConsultar
         CargarDatosAlosControles
     Case sghEliminar
         CargarDatosAlosControles
    End Select
    
    Select Case mi_Opcion
     Case sghAgregar
        Me.btnImprimir.Enabled = False
     Case sghModificar
        fraBusqueda.Enabled = False
        Me.btnImprimir.Enabled = True
        Me.chkPacienteNuevo.Enabled = False
     Case sghConsultar
        DeshabilitarControlesParaEdicion
        Me.btnAceptar.Visible = False
        Me.btnImprimir.Enabled = True
    Case sghEliminar
        DeshabilitarControlesParaEdicion
        Me.btnImprimir.Enabled = True
    End Select
 
End Sub
Sub DeshabilitarControlesParaEdicion()
    
    Me.chkPacienteNuevo.Enabled = False
    fraBusqueda.Enabled = False
    fraPacienteNuevo.Enabled = False
    
    HabilitarFrameOrigen False
    HabilitarFrameDestino False
   
    Me.ucPacientesDetalle1.DeshabilitarFrames

End Sub

Sub ValoresPorDefecto()

    'Obtiene la programacion medica
    Dim oDOProgramacion As New DOProgramacionMedica
    Dim oConexion As New Connection
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oDOProgramacion = mo_ReglasDeProgMedica.ProgramacionMedicaSeleccionarPorId(ml_IdProgramacion)
    
    mo_cmbIdEspecialidadMedico.BoundColumn = "IdEspecialidad"
    mo_cmbIdEspecialidadMedico.ListField = "DescripcionLarga"
    Dim rsEspecialidad As New Recordset
    Set rsEspecialidad = mo_AdminServiciosHosp.EspecialidadesSeleccionarporMedico(ml_IdMedico, oConexion)
    Set mo_cmbIdEspecialidadMedico.RowSource = rsEspecialidad
    
    mo_cmbIdEspecialidadMedico.BoundText = oDOProgramacion.IdEspecialidad
    mo_Formulario.HabilitarDeshabilitar cmbIdEspecialidadMedico, False
    
    cmbIdEspecialidadMedico_LostFocus
    
    Set mo_Especialidad = mo_AdminServiciosHosp.EspecialidadesSeleccionarPorId(Val(mo_cmbIdEspecialidadMedico.BoundText))
    
    mo_cmbIdServicio.BoundColumn = "IdServicio"
    mo_cmbIdServicio.ListField = "DescripcionLarga"
    Set rsServicio = mo_AdminServiciosHosp.ServiciosSeleccionarConsultoriosPorEspecialidaddebb(Val(mo_cmbIdEspecialidadMedico.BoundText), sghFiltraAnuladosYactivos, oConexion)
    Set mo_cmbIdServicio.RowSource = rsServicio
    
    mo_cmbIdServicio.BoundText = oDOProgramacion.IdServicio
    mo_Formulario.HabilitarDeshabilitar cmbIdServicio, False
    
    If mo_AdminServiciosComunes.MensajeError <> "" Then
        MsgBox mo_AdminServiciosComunes.MensajeError, vbInformation, Me.Caption
    End If
   
    mo_cmbIdViasAdmision.BoundText = 10 'D = Domicilio
    Me.txtMedico = DevuelveNombreMedicoPlanilla(Me.idMedico)
    
        
    Me.txtFechaIngreso.Text = mda_FechaIngreso
    Me.txtHoraInicio = ms_HoraInicio
    Me.txtHoraFin = ms_HoraFin
    mo_Formulario.HabilitarDeshabilitar txtHoraInicio, False
    mo_Formulario.HabilitarDeshabilitar txtHoraFin, False

    Me.ucPacientesDetalle1.ConfigurarValoresPorDefecto
    
    chkNuevoCarne.Value = vbUnchecked
    chkNuevoFolder.Value = vbUnchecked
    chkNuevoCarne.Visible = False
    chkNuevoFolder.Visible = False
    chkDuplicadoCarne.Visible = True
    
    oConexion.Close
    Set oConexion = Nothing
    Set oDOProgramacion = Nothing
End Sub
'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla CuentasAtencion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
    If mo_lbCargaTablasUnaVez = True Then
        '
        Me.ucCitasLista11.Inicializar
        Me.ucCitasLista11.Visible = False
        '
        lbCargaTablasUnaVez = False
        
        InicilizarParametros
        
        Me.ucPacientesDetalle1.Inicializar
        '
        UcPacientesSunasa1.Inicializar
        UcPacientesSunasa1.YaNoTieneSeguro
        '
        ucMensajeParpadeando2.MensajeDeTexto = "Cita Adicional"
        '
        '
        CargarComboBoxes
        '
        With grdPacientesEncontrados
            .Left = 240
            .Top = 780
            .Width = 11700 '11775
            .Height = 4455
        End With
        '
        If wxParametro282 = "S" Then
           Me.txtFichaFamiliar.Visible = True
           Me.lblFichaFamiliar.Visible = True
        End If
        '
        lbBuscaDNIenReniec = IIf(wxParametro296 = "S", True, False)
        If lbBuscaDNIenReniec = True Then
           mo_Reniec.SeAccesaAlaWebDesdeGalenhos = True
           mo_Reniec.Inicializar
        End If
        '
        If wxParametro302 = "S" Then
           Me.chkBuscarEnSIS.Visible = True
           UcSISafiliacion1.Visible = True
           UcSISafiliacion1.Inicializar
        End If
        '
        mo_Apariencia.ConfigurarFilasBiColores Me.grdPacientesEncontrados, sighEntidades.GrillaConFilasBicolor
        '
        Dim lcMensajeLicencia As String, mo_sighProxies As New SIGHProxies.Procesos
        lbTieneLicenciaTerapias = True
        Set mo_sighProxies = Nothing
        '
        
    End If
    '
    SiempreCargaPorMovimiento
    
End Sub

Sub SiempreCargaPorMovimiento()
    If mo_lbNuevoMovimiento = True Then
        
        lbCargaUnaVezVEntana = True
        '
        ucEPS1.Visible = False
        
        '
        tabAdmision.Top = 1920 '1830
        tabAdmision.Height = 6945
        If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE Then
           tabAdmision.Top = 0
           tabAdmision.Height = tabAdmision.Height + 1930
           '
           UcSISafiliacion1.Visible = False
        End If
        chkBuscarEnSIS.Value = 0
        '
        mo_lbNuevoMovimiento = False
        lbYaSeTransfirioHCdeUnServicioAotro = False
        mb_FormLoading = True
        mo_cmbIdTipoServicio.BoundText = ml_TipoServicio
        fraBusqueda.Enabled = False
        btnImprimeFichaSIS.Visible = False
        mo_Formulario.HabilitarDeshabilitar Me.FraGeneraCita, True
        Me.grdPacientesEncontrados.Visible = False
        Me.FraGeneraCita.Enabled = True
        Select Case mi_Opcion
        Case sghAgregar
            Me.Caption = "Agrega Admisión de CE"
            fraBusqueda.Enabled = True
            fraBusqueda.Visible = True
            If sighEntidades.Parametro583valorInt = "1" Then
               Me.optCweb.Enabled = True
            Else
               Me.optCweb.Enabled = False
            End If
            Me.optCNorma.Value = True

        Case sghModificar
            Me.Caption = "Modifica Admisión de CE"
            
        Case sghConsultar
            Me.Caption = "Consulta Admisión de CE"
        Case sghEliminar
            Me.Caption = "Elimina Admisión de CE"
        End Select
        '
        Me.UcPacientesSunasa1.idSunasaPacienteHistorico_idPaciente_ConValorCero
        Me.UcPacientesSunasa1.LimpiarDatos
        Me.UcPacientesSunasa1.PaisTitularDefault
        Me.UcPacientesSunasa1.YaNoTieneSeguro
        Me.UcPacientesSunasa1.Opcion = mi_Opcion
        '
        LimpiaTodosControles
        ConfiguraTABSsegunPermisosDelUsuario
        CargarDatosAlFormulario
        '
        If lbUsuarioTrabajaCitasPorTelefono = True Then
           Me.optCtelefono.Value = True
        End If
        '
        mo_Formulario.ConfigurarTipoLetra "Tahoma", "11", Me
        mo_Apariencia.ConfigurarFilasBiColores grdAnteriores, sighEntidades.GrillaConFilasBicolor
        grdAnteriores.Caption = ""
        If mi_Opcion = sghAgregar Then
           btnAceptar.Enabled = True
        End If
        btnAceptar.Visible = True
        ucMensajeParpadeando1.Visible = False
        Me.grdPacientesEncontrados.Visible = False
        Set Me.grdAnteriores.DataSource = Nothing
        '
        ucPacientesDetalle1.MarcoCheckPacienteNuevo = False
        ucPacientesDetalle1.Opcion = mi_Opcion
        lc_AntecedentePersonal = ""
        '
        lbElConsultorioNoEnviaMensajeTextoCelular = False
        lbElMedicoNOregistraFUA = "N"
        lbElConsultorioNoCobraApagantes = False
        If rsServicio.RecordCount > 0 And Val(mo_cmbIdServicio.BoundText) > 0 Then
           rsServicio.MoveFirst
           rsServicio.Find "idServicio=" & mo_cmbIdServicio.BoundText
           If Not rsServicio.EOF Then
              If rsServicio.Fields!UsaGalenHos = True Then
                 lbElMedicoNOregistraFUA = "S"
              End If
              If rsServicio.Fields!UsaFUA <> True Or IsNull(rsServicio.Fields!UsaFUA) Then
                   wxParametro302 = "N"
              Else
                 wxParametro302 = lcBuscaParametro.SeleccionaFilaParametro(302)
              End If
              If rsServicio!NoUsaMensajeTexto = 1 Then
                 lbElConsultorioNoEnviaMensajeTextoCelular = True
              End If
              '
              If Not IsNull(rsServicio!CostoCeroCE) Then
                  If rsServicio!CostoCeroCE = "S" Then
                     lbElConsultorioNoCobraApagantes = True
                  End If
              End If
              
           End If
        End If
        'wxParametro302 = lcBuscaParametro.SeleccionaFilaParametro(302) ' borrar frank
        If wxParametro302 = "S" Then
           Me.ucSISfuaCodPrestacion1.CodigoPrestacion = ""
           Me.chkBuscarEnSIS.Visible = True
           UcSISafiliacion1.Visible = True
          ' If mi_Opcion = sghModificar Then
              Me.ucSISfuaCodPrestacion1.Visible = True
          ' End If
        Else
           Me.chkBuscarEnSIS.Visible = False
           UcSISafiliacion1.Visible = False
           Me.ucSISfuaCodPrestacion1.Visible = False
        End If
        HaceVisibleOnoBotonFUA
        '
    End If
    
End Sub

Sub LimpiaTodosControles()
    If mi_Opcion = sghAgregar Then
            Me.ucPacientesDetalle1.LimpiarDatosDePaciente wxParametro211, ldFechaActualServidor
            Me.ucPacientesDetalle1.HabilitarFrames
            mo_Pacientes.idPaciente = 0
            '
            fraPacienteNuevo.Enabled = True
            Me.chkPacienteNuevo.Enabled = True
            '
            chkPacienteNuevo.Value = 0
            '
            Me.idCuentaAtencion = 0
            Me.idPaciente = 0
            mo_cmbIdViasAdmision.BoundText = ""
            mo_cmbIdTipoReferenciaOrigen.BoundText = ""
            Me.txtIdEstablecimientoOrigen.Tag = ""
            mo_cmbIdFormaPago.BoundText = ""
            mo_cmbIdFuentesFinanciamiento.BoundText = ""
            '
            txtNroHistoriaBusqueda.Text = ""
            txtApellidoPaternoBusqueda.Text = ""
            txtApellidoMaternoBusqueda.Text = ""
            txtPrimerNombreBusqueda.Text = ""
            txtSegundoNombreBusqueda.Text = ""
            txtNroDNIBusqueda.Text = ""
            lnIdDistritoSIS = 0: lnIdSexoSIS = 0: ldFechaNacimientoSIS = 0: lcSnombreSIS = "": lnIdPlanSIS = 0
            UcSISafiliacion1.InabilitaControles True: lcDniSIS = "": lnAfiliacionSIS1 = ""
            lnAfiliacionSIS2 = "": lnAfiliacionSIS3 = "": lnAfiliacionSIS4 = 0: lcSIScodigo = ""
            Me.ucSISfuaCodPrestacion1.Visible = False
            If wxParametro302 = "S" Then
               Me.ucSISfuaCodPrestacion1.CodigoPrestacion = ""
            End If
            txtNroOrdenPago.Text = ""
            Me.txtNroCuenta.Text = ""
            txtNboleta.Text = ""
            txtNserie.Text = ""
            Me.txtFichaFamiliar.Text = ""
            '
            lblBoleta.Visible = False
            txtNserie.Visible = False
            txtNboleta.Visible = False
            '
            mo_DOFacturacionPaquetes.IdComprobantePago = 0
            mo_DOFacturacionPaquetes.IdOrdenPago = 0
            mo_DOFacturacionPaquetes.idProducto = 0
            '
            btnImprimeFiliacion.Enabled = False
            cmbServicioReferenciaO.Text = "": txtDxReferencia.Text = "": txtDxReferencia.Tag = "": lblDxReferencia1.Text = ""      'debb-21/06/2016
    End If
    '
    lcHistoriaYpaciente = ""
    txtReferenciaO.Text = ""
    '
    txtMedicoRef.Text = "": Me.cmbMedicoRef.Text = ""  'franklin 2017
End Sub

Sub ConfiguraTABSsegunPermisosDelUsuario()
    Dim ms_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
    Dim oRsPermisosTabs As New Recordset
    lbUsuarioTrabajaCitasPorTelefono = False
    Me.tabAdmision.TabVisible(0) = False
    Me.tabAdmision.TabVisible(1) = False
    lbUsuarioAutorizadoAregistrarCitasRepetidas = False
    Set oRsPermisosTabs = ms_ReglasSeguridad.UsuariosRolesSeleccionarPermisosTodos(ml_idUsuario)
    If oRsPermisosTabs.RecordCount > 0 Then
       Do While Not oRsPermisosTabs.EOF
          Select Case oRsPermisosTabs.Fields!IdPermiso
          Case 359    'Admisión CE - Ver TAB  Paciente (F10)
               Me.tabAdmision.TabVisible(0) = True
          Case 360    'Admisión CE - Ver TAB  Cita (F11)
               Me.tabAdmision.TabVisible(1) = True
          Case 361    'Admisión CE - Ver TAB  Alta (F12)
          Case 365
               lbUsuarioAutorizadoAregistrarCitasRepetidas = True
          Case 410
               lbUsuarioTrabajaCitasPorTelefono = True
          End Select
          oRsPermisosTabs.MoveNext
       Loop
    End If
    Set oRsPermisosTabs = Nothing
    Set ms_ReglasSeguridad = Nothing
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla CuentasAtencion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Activate()
   SiempreCargaPorMovimiento
   If mi_Opcion <> sghAgregar Then
       If Not mb_ExistenDatos Then
           Me.Visible = False
           LimpiarVariablesDeMemoria
       End If
       If sighEntidades.EsFecha(txtFechaIngreso.Text, "DD/MM/AAAA") Then
            If CDate(txtFechaIngreso.Text) < ldFechaActualServidor And (mi_Opcion = sghModificar Or mi_Opcion = sghEliminar) And mo_lnIdTablaLISTBARITEMS <> 103 Then
                 MsgBox "No puede Modificar/Eliminar CITAS de días menores a: " & ldFechaActualServidor, vbInformation, Me.Caption
                 
                 LimpiarVariablesDeMemoria
                 Me.Visible = False
            Else
            
            End If
       Else
            MsgBox "Hubo problemas al cargar la ATENCION" + Chr(13) + Chr(13) + "Salga del Sistema a Windows y vuelva a ingresar", vbInformation, Me.Caption
            Me.Visible = False
            LimpiarVariablesDeMemoria
       End If
       If lbElMedicoNOregistraFUA = "S" And mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE Then
            MsgBox "El CONSULTORIO se configuró para que no se registre ATENCION DEL PACIENTE", vbInformation, Me.Caption
            
            LimpiarVariablesDeMemoria
            Me.Visible = False
          
       End If
   Else
       btnBuscaHistoricos.Visible = False
       If CDate(txtFechaIngreso.Text) < ldFechaActualServidor Then
            MsgBox "No puede registrar CITAS de días menores a: " & ldFechaActualServidor, vbInformation, Me.Caption
            Me.Visible = False
            LimpiarVariablesDeMemoria
        End If
        If Me.ucCitasLista11.Visible = True And TabIngreso.Tab = 1 Then
           Me.ucCitasLista11.NOCargaDesdeCitas = True
        End If
   
   End If
   fraDatosCita.Caption = "Cita N° " & ms_NroCola
   ucMensajeParpadeando2.Visible = mo_lbEsCitaAdicional
   If mb_FormLoading Then
        On Error Resume Next
        Select Case mi_Opcion
        Case sghAgregar
            tabAdmision.Tab = 0
            If lbCargaUnaVezVEntana = True Then
                lbCargaUnaVezVEntana = False
                Select Case WxDEFAULT_BUSQ_CE
                Case sghDefaultVentana.sighApellidoPaterno
                     txtApellidoPaternoBusqueda.SetFocus
                Case sghDefaultVentana.sighDNI
                     txtNroDNIBusqueda.SetFocus
                Case sghDefaultVentana.sighHistoria
                     txtNroHistoriaBusqueda.SetFocus
                End Select
            End If
        Case sghModificar
            Me.ucPacientesDetalle1.SetFocusOnApellidoPaterno
            If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE Then
               Dim oConexion As New Connection
               oConexion.Open sighEntidades.CadenaConexion
               oConexion.CursorLocation = adUseClient
               MuestraCitasAnteriores oConexion, False
               oConexion.Close
               Set oConexion = Nothing
               AdministrarKeyPreview vbKeyF12
            Else
               AdministrarKeyPreview vbKeyF11
            End If
        Case sghConsultar
        Case sghEliminar
        End Select
        '
        mb_FormLoading = False
        lbCargaUnaSolaVez = True
    End If
    'franklin 2017
    If mi_nroHistoriaCitadoXmedico > 0 Then
        Me.tabAdmision.TabVisible(0) = True
        Me.tabAdmision.TabVisible(1) = True
        Me.txtNroHistoriaBusqueda.Text = mi_nroHistoriaCitadoXmedico
        btnBuscarPaciente_Click
        Me.tabAdmision.Tab = 1
        ml_idUsuario = sighEntidades.Usuario
        lcCodigoEstablecimientoAdscripcionSIS = ml_lcCodigoEstablecimientoAdscripcionSISxMedico
        mo_cmbIdViasAdmision.BoundText = ml_cmbIdViasAdmisionXmedico
        mo_cmbIdTipoReferenciaOrigen.BoundText = ml_cmbIdTipoReferenciaOrigenXmedico
        txtReferenciaO.Text = ml_txtReferenciaOXmedico
        '
        Dim oDoEstablecimiento As New DOEstablecimiento
        Set oDoEstablecimiento = mo_AdminServiciosComunes.EstablecimientosSeleccionarPorId(Val(ml_txtIdEstablecimientoOrigenXmedico))
        If Not oDoEstablecimiento Is Nothing Then
            txtIdEstablecimientoOrigen.Tag = oDoEstablecimiento.IdEstablecimiento
            txtIdEstablecimientoOrigen.Text = oDoEstablecimiento.Codigo
            txtNombreOrigenReferencia.Text = oDoEstablecimiento.nombre
        End If
        Set oDoEstablecimiento = Nothing
        '
        Dim lcDxCodigo As String, lcDx As String, oConexion1 As New Connection
        PVcomboBoxUbicaPosicion ml_cmbServicioReferenciaOXmedico, cmbServicioReferenciaO
        txtDxReferencia.Tag = ml_txtDxReferenciaXmedico
        sighEntidades.AbreConexionSIGH oConexion1
        mo_AdminServiciosComunes.DiagnosticosSeleccionarPorIdDevuelveDescripcion Val(txtDxReferencia.Tag), _
                                                                                 oConexion1, lcDxCodigo, lcDx
        txtDxReferencia.Text = lcDxCodigo
        lblDxReferencia1.Text = lcDx
        oConexion1.Close
        Set oConexion1 = Nothing
        '
        Me.txtMedicoRef.Text = ml_txtMedicoRefXMedico
        BuscaMedicoRerencia ml_txtMedicoRefXMedico
        '
        mo_cmbIdFuentesFinanciamiento.BoundText = Trim(Str(ml_idFuenteFinanciamientoCitadoXmedico))
        mo_cmbIdFormaPago.BoundText = Trim(Str(ml_idFormaPagoCitadoXmedico))
        Frame5.BackColor = vbCyan
        fraBusqueda.Visible = False
        grdAnteriores.Visible = False
        fraDatosCita.BackColor = vbCyan
        FraGeneraCita.BackColor = vbCyan
        If txtReferenciaO.Text <> "" Then fraDatosReferenciaOrigen.BackColor = vbCyan
        Me.Caption = "ES UNA CITA PARA TERAPIAS"
        lblNroAtencion.FontBold = True
        btnAceptar_Click
    ElseIf Me.ucCitasLista11.Visible = True And TabIngreso.Tab = 1 Then
        Dim oConexion987 As New Connection
        sighEntidades.AbreConexionSIGH oConexion987
        MuestraCitasAnteriores oConexion987, True
        oConexion987.Close
        Set oConexion987 = Nothing
    End If

End Sub
Sub AdministrarKeyPreview(KeyCode As Integer)
    
    Select Case KeyCode
    'Case vbKeyEscape
    '    btnCancelar_Click
    Case vbKeyF2
        btnAceptar_Click
    Case vbKeyF3
        btnImprimir_Click
    Case vbKeyF6
         btnBuscarPaciente_Click
     Case vbKeyF7
         Me.tabAdmision.Tab = 0
         Me.ucPacientesDetalle1.SetPestaniaTabPaciente 0
         On Error Resume Next
         Me.ucPacientesDetalle1.SetFocusOnDepartamentoDomicilio
     Case vbKeyF8
         Me.tabAdmision.Tab = 0
         Me.ucPacientesDetalle1.SetPestaniaTabPaciente 1
         On Error Resume Next
         Me.ucPacientesDetalle1.SetFocusOnDepartamentoProcedencia
     Case vbKeyF9
         Me.tabAdmision.Tab = 0
         On Error Resume Next
         Me.ucPacientesDetalle1.SetPestaniaTabPaciente 2
         Me.ucPacientesDetalle1.SetFocusOnDepartamentoNacimiento
     Case vbKeyF10
         Me.tabAdmision.Tab = 0
         On Error Resume Next
         Me.ucPacientesDetalle1.SetFocusOnApellidoPaterno
     Case vbKeyF11
         Me.tabAdmision.Tab = 1
         
         On Error Resume Next
         cmbFuenteFinanciamiento.SetFocus
         
    End Select
       
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   AdministrarKeyPreview KeyCode
End Sub

Sub ImprimeFormatoSIS(lnIdAtencion As Long, lnIdUsuarioSistema As Long)
    Dim oImprimeSIS As New RptHistoriaClinicaCE
    oImprimeSIS.ImprimeFormatoSIS lnIdAtencion, lnIdUsuarioSistema, 0
    Set oImprimeSIS = Nothing
End Sub

Private Sub btnAceptar_Click()
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   Dim oConexion As New Connection
   oConexion.CommandTimeout = 300
   oConexion.CursorLocation = adUseClient
   oConexion.Open sighEntidades.CadenaConexion
   
   Select Case mi_Opcion
   Case sghAgregar
       
       GuardaFua = "F" 'HRA 10/12/2020 Cambio 47
       
       If ValidarDatosObligatorios() Then
            CargaDatosAlObjetosDeDatos
            If ValidarReglas() Then
                'mgaray201503
                Dim lNroOrdenPago As Long
                If AgregarDatos() Then
                    If Val(txtNroOrdenPago.Text) > 0 Then
                       lcApP = "N° Orden Pago: " & txtNroOrdenPago.Text
                       'mgaray201503
                       lNroOrdenPago = txtNroOrdenPago.Text
                       '
                       If lbElConsultorioNoCobraApagantes = True Then
                           mo_AdminServiciosComunes.AnularCptGeneradasEnCitas mo_Atenciones.idCuentaAtencion
                       End If

                    Else
                       lcApP = ""
                    End If
                    
                    '
                    mo_AdminAdmision.CitasFormaCitaActualiza mo_Atenciones.idCuentaAtencion, _
                                                           IIf(Me.optCNorma.Value = True, "N", _
                                                           IIf(Me.optCweb.Value = True, "W", "T")), _
                                                           oConexion
                    '
                    If lbTieneLicenciaParaMensajeAcelulares = True And lbElConsultorioNoEnviaMensajeTextoCelular = False Then
                        Dim oMensajeCelular As New SIGHProxies.Procesos
                        Dim lcMensajeCelular As String
                        lcMensajeCelular = "Cita para " & mo_Atenciones.FechaIngreso & " " & mo_Atenciones.HoraIngreso & _
                                           IIf(lcApP = "", "", " (" & lcApP & ")") & _
                                           " (N° Cuenta: " & mo_Atenciones.idCuentaAtencion & ") (Consultorio: " & _
                                            Trim(Mid(cmbIdServicio.Text, InStr(cmbIdServicio.Text, "=") + 1, 100)) & _
                                            ")  en " & wxParametro205
                        oMensajeCelular.MensajeCelularEnviar mo_Pacientes, mo_Atenciones.idCuentaAtencion, lcMensajeCelular, _
                                        "CITAS", oConexion
                        Set oMensajeCelular = Nothing
                    End If
                    '
                    'mgaray201503
                    Dim bAbrirModuloCaja As Boolean
                    bAbrirModuloCaja = False
                    If bEsCajero = True And lcApP <> "" Then
                        bAbrirModuloCaja = True
                    End If
                    ms_NombrePaciente = mo_paciente.ApellidoPaterno + " " + mo_paciente.ApellidoMaterno + " " + mo_paciente.PrimerNombre
                    Me.idAtencion = mo_Atenciones.idAtencion
                    Me.txtNroCuenta = mo_Atenciones.idCuentaAtencion
                    lblEstadoCta.Caption = "Abierta"
                    Me.btnImprimir.Enabled = True
                    Me.btnAceptar.Enabled = False
                    If mi_nroHistoriaCitadoXmedico = 0 Then
                      MsgBox "Los datos se agregaron correctamente para la Historia Nª:  " & _
                           HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(mo_Pacientes.NroHistoriaClinica)), False) & _
                      Chr(13) & Chr(13) & "N° Cuenta: " & Me.txtNroCuenta.Text & Chr(13) & Chr(13) & lcApP, vbInformation, Me.Caption
                    End If
                    mo_AdminArchivoClinico.generadorNroHistoriaClinicaActualizaNroAutomaticoDeHistoriaClinica oConexion
                    
                    If txtNroCuenta.Text <> "" Then
                        lbImpresionCuenta = True
                        ImprimePreCuenta
                        If chkPacienteNuevo.Value = 1 Then
                            btnImprimeFiliacion.Enabled = True
                        Else
                            Me.Visible = False
                            LimpiarVariablesDeMemoria
                        End If
                    End If
                    ActualizaCitas_Atencion oConexion   'franklin 2017
                    'mgaray201503
                    If bAbrirModuloCaja Then
                        If Principal.bAbrioCaja = True Then
                            FrmGestionCajaDesdeServicios.lIdUsuarioAuditoria = ml_idUsuario
                            FrmGestionCajaDesdeServicios.lcNombrePc = mo_lcNombrePc
                            FrmGestionCajaDesdeServicios.lNumeroOrden = lNroOrdenPago
                            FrmGestionCajaDesdeServicios.CerrarAlGuardar = True
'                            FrmGestionCajaDesdeServicios.ucGestionCaja1.ActivarOrdenExistenteFS 1
'                            FrmGestionCajaDesdeServicios.ucGestionCaja1.AsignarNroOrden CStr(lNroOrdenPago)
                            FrmGestionCajaDesdeServicios.Show 1
                        Else
                            If mi_nroHistoriaCitadoXmedico = 0 Then
                               MsgBox "No se ha aperturado caja, para poder realizar cobros proceda a aperturar caja desde el modulo de Gestion deCajas", vbInformation, Me.Caption
                            End If
                        End If
                    End If
                    
                    If wxParametro302 = "S" And mo_Atenciones.IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS And lbElMedicoNOregistraFUA = "S" Then
                     
                        btnImprimeFichaSIS_Click
                    ElseIf mo_Atenciones.IdFormaPago = sghTipoFinanciamiento.sghSis Then
                        If Val(wxParametro208) = 1910 Then  'sullana
                            lbImpresionCuenta = False
                            ImprimePreCuenta
                        End If
                    End If
                    lcApP = ""
               Else
                    ms_NombrePaciente = ""
                    MsgBox "No se pudo agregar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
       
   Case sghModificar
     
       If ValidarDatosObligatorios() Then
            CargaDatosAlObjetosDeDatos
            If ValidarReglas() Then
               If ModificarDatos() Then
                  mo_AdminAdmision.CitasFormaCitaActualiza mo_Atenciones.idCuentaAtencion, _
                                                           IIf(Me.optCNorma.Value = True, "N", _
                                                           IIf(Me.optCweb.Value = True, "W", "T")), _
                                                           oConexion
                   
                   ActualizaCitas_Atencion oConexion   'franklin 2017
                   ms_NombrePaciente = mo_paciente.ApellidoPaterno + " " + mo_paciente.ApellidoMaterno + " " + mo_paciente.PrimerNombre
                   MsgBox " Los datos se modificaron correctamente, para la Cuenta N°: " & Me.txtNroCuenta.Text & DevuelveNroRecetasGeneradas, vbInformation, Me.Caption
                   PagaBoletaPendientePorConsulta
                   VerSiTieneServicioAutomaticoPorEstancia oConexion
                   If txtNroCuenta.Text <> "" And mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroCitaCE Then
                       lbImpresionCuenta = True
                       ImprimePreCuenta
                   End If
                   If wxParametro302 = "S" And mo_Atenciones.IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                         btnImprimeFichaSIS_Click
                   ElseIf mo_Atenciones.IdFormaPago = sghTipoFinanciamiento.sghSis Then
                        lbImpresionCuenta = False
                        ImprimePreCuenta
                   End If
                   Me.Visible = False
                   LimpiarVariablesDeMemoria
               Else
                   ms_NombrePaciente = ""
                   MsgBox "No se pudo modificar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
            CargaDatosAlObjetosDeDatos
            If ValidarReglas() Then
               If EliminarDatos(oConexion) Then
                    ActualizaCitas_Atencion oConexion   'franklin 2017
                    MsgBox " Los datos se eliminaron correctamente, para la Cuenta N°: " & Me.txtNroCuenta.Text, vbInformation, Me.Caption
                    Me.Visible = False
                    LimpiarVariablesDeMemoria
               Else
                    ms_NombrePaciente = ""
                    MsgBox "No se pudo eliminar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
               End If
           End If
   End Select
   oConexion.Close
   Set oConexion = Nothing
End Sub


Function VerSiTieneServicioAutomaticoPorEstancia(oConexion As Connection) As String
    Dim oRsTmp As New ADODB.Recordset
    txtNroOrdenPago.Text = ""
    VerSiTieneServicioAutomaticoPorEstancia = ""
    Set oRsTmp = mo_AdminFacturacion.FactOrdenServicioPagosPorIdAtencion(mo_Atenciones.idAtencion, oConexion)
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.Filter = "idPuntoCarga=6"
    End If
    If oRsTmp.RecordCount > 0 Then
       VerSiTieneServicioAutomaticoPorEstancia = Chr(13) & "(Ord.Pago)= "
       oRsTmp.MoveFirst
       txtNroOrdenPago.Text = oRsTmp.Fields!IdOrdenPago
       Do While Not oRsTmp.EOF
          VerSiTieneServicioAutomaticoPorEstancia = VerSiTieneServicioAutomaticoPorEstancia & Trim(Str(oRsTmp.Fields!IdOrdenPago)) & " , "
          oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
End Function


'El paciente PAGA su CONSULTA-BRIGADAS
Sub PagaBoletaPendientePorConsulta()
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    If wxParametro208 = "1" Then
       Dim lcSql As String
       Dim oRsTmp  As New ADODB.Recordset
       Dim oConexion As New ADODB.Connection
       Dim oDOFactOrdenServicio As DOFactOrdenServicio
       Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
       Dim lnOrdenes As Long
       Dim lnIdPuntoCarga As Long
       Dim mo_DoAtencion As New DOAtencion
       Dim lcFechaIngreso As String
       Dim oDOCajaComprobantesPago As New DOCajaComprobantesPago
       Dim mo_DOComprobantePago As New DOCajaComprobantesPago
       Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
       Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
       Dim mo_DoPaciente As New doPaciente
       Dim lcNombres As String
       Dim IdTipoGenHistoriaClinica As Long
       Dim lcRazonSocial As String
       Dim oCajaNroDocumento As New DOCajaNroDocumento
       Dim lcNroDocumento As String
       Dim lcNroSerie  As String
       Dim lcNroHistoria As String
       Dim lnTotal As Double
       Dim lbAgregarDatos  As Boolean
       Dim mo_doCajaGestion As New DOCajaGestion
       Dim mo_DOFactOrdenServicio As New DOFactOrdenServicio
       Dim oDllFactUCGestionCaja As New SighFacturacion.dllFactUCGestionCaja

       oConexion.CommandTimeout = 130
       oConexion.Open sighEntidades.CadenaConexion
       Set oRsTmp = mo_AdminFacturacion.FactOrdenServicioSeleccionarSoloAdmisionCE(mo_Atenciones.idAtencion)
       If oRsTmp.RecordCount > 0 Then
          ucFacturacionProductos.TipoProducto = sghServicio
          ucFacturacionProductos.idUsuario = ml_idUsuario
          ucFacturacionProductos.Inicializar
          ucFacturacionProductos.EstadosFacturacion = "1,3"    'Registrados y pendientes de pago
          ucFacturacionProductos.TiposFinanciamiento = "1,5,9"
       
          oRsTmp.MoveFirst
          Do While Not oRsTmp.EOF
                '++++++++++++++++++++++++++++++++++Carga Datos
                lnOrdenes = oRsTmp.Fields!IdOrden
                'Carga datos de la orden
                Set oDOFactOrdenServicio = mo_ReglasFacturacion.FactOrdenServicioSeleccionarPorId(lnOrdenes)
                Set mo_DOFactOrdenServicio = oDOFactOrdenServicio
                'Cargar datos del paciente y de la atencion
                Set mo_DoAtencion = mo_AdminAdmision.AtencionesSeleccionarPorId(oDOFactOrdenServicio.idAtencion, oConexion)
                Set ucFacturacionProductos.Atencion = mo_DoAtencion
                lcFechaIngreso = mo_DoAtencion.FechaIngreso
                Set mo_DOComprobantePago = mo_AdminCaja.ComprobantePagoSeleccionarPorId(oDOFactOrdenServicio.IdComprobantePago, oConexion)
                Set mo_DoPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(mo_DoAtencion.idPaciente, oConexion)
                If Not mo_DoPaciente Is Nothing Then
                    lcNombres = mo_DoPaciente.ApellidoPaterno + " " + mo_DoPaciente.ApellidoMaterno + " " + mo_DoPaciente.PrimerNombre
                    IdTipoGenHistoriaClinica = mo_DoPaciente.idTipoNumeracion
                    lcNroHistoria = mo_DoPaciente.NroHistoriaClinica
                    lcRazonSocial = lcNombres
                    ml_IdPaciente = mo_DoPaciente.idPaciente
                End If
                'Cargar datos de los servicios
                ucFacturacionProductos.LimpiarGrilla
                ucFacturacionProductos.IdOrden = lnOrdenes
                ucFacturacionProductos.CargaProductosPorIdOrden
                lnTotal = ucFacturacionProductos.DevuelveTotalPagar
                '++++++++++++++++++++++++++++++++++Carga ultimo Comprobante ANTES DE GRABAR
                Set oCajaNroDocumento = mo_AdminCaja.NroDocumentoSeleccionarPorIdCajaYTipoComprobante(5, 3)
                lcNroSerie = Trim(oCajaNroDocumento.nroSerie)
                lcNroDocumento = Trim(oCajaNroDocumento.nrodocumento)
                
                '++++++++++++++++++++++++++++++++++Carga Datos ANTES DE GRABAR
                mo_doCajaGestion.IdGestionCaja = 108  '
                mo_doCajaGestion.IdCajero = 738       'daniel barrantes
                mo_doCajaGestion.IdCaja = 5           'Caja nueva
                Set mo_DOComprobantePago = New DOCajaComprobantesPago
                With mo_DOComprobantePago
                    .IdTipoComprobante = 3
                    .nroSerie = Trim(lcNroSerie)
                    .nrodocumento = Trim(lcNroDocumento)
                    .idCuentaAtencion = mo_DoAtencion.idCuentaAtencion
                    .razonSocial = lcRazonSocial
                    .Observaciones = ""
                    .IdGestionCaja = mo_doCajaGestion.IdGestionCaja
                    .IdUsuarioAuditoria = ml_idUsuario
                    .ruc = ""
                    .Subtotal = lnTotal
                    .IGV = 0
                    .Total = lnTotal  'CCur(txtTotal.Text)
                    .FechaCobranza = Now
                    .IdComprobantePago = 0
                    .IdTipoPago = 1 'Orden de pago
                    .idPaciente = ml_IdPaciente
                    .IdFormaPago = 1  'contado
                    .idFarmacia = 0   'otros
                End With
                '++++++++++++++++++++++++++++++++++Graba datos
                'lbAgregarDatos = mo_AdminCaja.CajaComprobantePagoServicioAgregar(mo_DOComprobantePago, mo_doCajaGestion, mo_DOFactOrdenServicio, ucFacturacionProductos.FacturacionProductos, ml_idUsuario, mo_Atenciones, 6, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, 0)
                lbAgregarDatos = oDllFactUCGestionCaja.CajaComprobantePagoServicioAgregar(mo_DOComprobantePago, mo_doCajaGestion, mo_DOFactOrdenServicio, ucFacturacionProductos.FacturacionProductos, ml_idUsuario, mo_Atenciones, 6, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, 0, False)
                '
                If lbAgregarDatos = True Then
                   MsgBox "Se agregó su BOLETA DE PAGO correctamente"
                Else
                   MsgBox "Hubo problemas para registrar la BOLETA DE PAGO"
                   Exit Do
                End If
             
             oRsTmp.MoveNext
          Loop
       End If
       oRsTmp.Close
       oConexion.Close
       Set oRsTmp = Nothing
       Set oConexion = Nothing
       Set oDllFactUCGestionCaja = Nothing
    End If
End Sub

Private Sub btnCancelar_Click()
   Me.Visible = False
   'mgaray20141003
    lcDniSIS = ""
    lcApM = ""
    lcPnom = ""
   LimpiarVariablesDeMemoria
End Sub

Function ValidarDatosObligatorios() As Boolean
   Dim sMensaje As String
   ValidarDatosObligatorios = False
         
    '---------------------------------------------------------------------------------
    '           VALIDA DATOS DE LA ATENCION
    '---------------------------------------------------------------------------------
   If Val(mo_cmbIdEspecialidadMedico.BoundText) = 0 Then
       sMensaje = sMensaje + "Ingrese el valor de la especialidad del médico" + Chr(13)
   End If
   If Val(mo_cmbIdServicio.BoundText) = 0 Then
       sMensaje = sMensaje + "Ingrese el valor del servicio de ingreso" + Chr(13)
   End If
   If Me.txtHoraInicio.Text = "" Then
       sMensaje = sMensaje + "ingrese el valor de la hora ingreso" + Chr(13)
   End If
   If Me.txtFechaIngreso.Text = sighEntidades.FECHA_VACIA_DMY Then
       sMensaje = sMensaje + "Ingrese el valor de la fecha de ingreso" + Chr(13)
   End If
   If Val(mo_cmbIdTipoServicio.BoundText) = 0 Then
       sMensaje = sMensaje + "Ingrese el valor del tipo de servicio" + Chr(13)
   End If
   If Val(Me.txtEdadEnDias.Text) = 0 Then
       sMensaje = sMensaje + "Ingrese el valor de la edad" + Chr(13)
       Me.tabAdmision.Tab = 0
   End If
   If (mi_Opcion = sghAgregar) Or (mi_Opcion = sghModificar) Then
        If Val(mo_cmbIdTipoConsulta.BoundText) = 0 Then
             sMensaje = sMensaje + "Seleccione el tipo de consulta" + Chr(13)
        End If
   End If
   If Me.cmbFuenteFinanciamiento.Text = "" Then
       sMensaje = sMensaje + "Elija el Plan de Atención" + Chr(13)
       Me.tabAdmision.Tab = 1
   End If
   If Val(mo_cmbIdFormaPago.BoundText) = 0 Then
      sMensaje = sMensaje + "Por favor elija el Tipo Financiamiento (Ficha 'Cita')" + Chr(13)
   End If
   
    '---------------------------------------------------------------------------------
    '           VALIDA DATOS DE PACIENTES
    '---------------------------------------------------------------------------------
    ms_MensajeError = ucPacientesDetalle1.ValidarDatosObligatorios(wxParametro282, wxParametro333)
    If ms_MensajeError <> "" Then
        If ucPacientesDetalle1.DevuelveEtnia = "" Then
          Me.tabAdmision.Tab = 0
          ucPacientesDetalle1.SetFocusEnEtnia
        ElseIf ucPacientesDetalle1.DevuelveIdioma = "" Then
          Me.tabAdmision.Tab = 0
          ucPacientesDetalle1.SetFocusEnIdioma
        End If
    End If
    'debb-21/06/2016 (inicio)
    If ValidaDatosObligatoriosReferencias(sMensaje) Then
    End If
'    If (mo_cmbIdViasAdmision.BoundText = "12" Or mo_cmbIdViasAdmision.BoundText = "13") Then
'        If Val(txtIdEstablecimientoOrigen.Text) = 0 Then
'           sMensaje = sMensaje + "Por favor elija el ORIGEN DE LA REFERENCIA (Ficha 'Cita')" + Chr(13)
'        End If
'        If txtReferenciaO.Text = "" Then
'           sMensaje = sMensaje + "Por favor ingrese el N° DE REFERENCIA (Ficha 'Cita')" + Chr(13)
'        End If
'        If cmbServicioReferenciaO.Text = "" Then
'           sMensaje = sMensaje + "Por favor debe elegir el SERVICIO DE LA REFERENCIA (Ficha 'Cita')" + Chr(13)
'        End If
'        'FRANKLIN 2017
'        If lcBuscaParametro.SeleccionaFilaParametro(516) = "S" And (txtMedicoRef.Text = "" Or Me.cmbMedicoRef.Text = "") Then
'           sMensaje = sMensaje + "Por favor debe ingresar: COLEGIATURA, APELLIDOS Y NOMBRES DEL MEDICO QUE REFIERE (Ficha 'Cita')" + Chr(13)
'        End If
'        If wxParametro580 = "S" And lblDxReferencia1.Text = "" Then
'           sMensaje = sMensaje + "Por favor debe ingresar: Dx QUE REFIERE (Ficha 'Cita')" + Chr(13)
'        End If
'    End If
    'debb-21/06/2016 (fin)
    If ms_MensajeError <> "" Then
       Me.tabAdmision.Tab = 0
       sMensaje = sMensaje + ms_MensajeError
    End If
    '
'    If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroCitaCE Then
'       If mo_ReglasSISgalenhos.ReglasDeConsistenciaSISestanOK(sghOpcionGalenHos.sghRegistroCitaCE, wxParametro302, _
'                                                           Val(mo_cmbIdFuentesFinanciamiento.BoundText), sMensaje, 0, _
'                                                           Me.ucSISfuaCodPrestacion1.CodigoPrestacion, lbElMedicoNOregistraFUA, _
'                                                           sghConsultaExterna, mi_Opcion, True, False, 0, "", "") = True Then
'       End If
'    End If
    '
    If Me.ucEPS1.Visible = True Then
       If Me.ucEPS1.ValidaDatosObligatorios = False Then
          sMensaje = sMensaje + Me.ucEPS1.MensajeError
       End If
    End If
    '
    If sMensaje <> "" Then
        MsgBox sMensaje, vbInformation, Me.Caption
        Exit Function
    End If
    ValidarDatosObligatorios = True

End Function
Function ValidaDatosObligatoriosReferencias(ByRef lcMensaje As String) As Boolean
    ValidaDatosObligatoriosReferencias = True
    Dim sMensaje As String
    sMensaje = ""
    If (mo_cmbIdViasAdmision.BoundText = "12" Or mo_cmbIdViasAdmision.BoundText = "13") Then
        If Val(txtIdEstablecimientoOrigen.Text) = 0 Then
           sMensaje = sMensaje + "Por favor elija el ORIGEN DE LA REFERENCIA (Ficha 'Cita')" + Chr(13)
        End If
        If txtReferenciaO.Text = "" Then
           sMensaje = sMensaje + "Por favor ingrese el N° DE REFERENCIA (Ficha 'Cita')" + Chr(13)
        End If
        If cmbServicioReferenciaO.Text = "" Then
           sMensaje = sMensaje + "Por favor debe elegir el SERVICIO DE LA REFERENCIA (Ficha 'Cita')" + Chr(13)
        End If
        'FRANKLIN 2017
        If lcBuscaParametro.SeleccionaFilaParametro(516) = "S" And (txtMedicoRef.Text = "" Or Me.cmbMedicoRef.Text = "") Then
           sMensaje = sMensaje + "Por favor debe ingresar: COLEGIATURA, APELLIDOS Y NOMBRES DEL MEDICO QUE REFIERE (Ficha 'Cita')" + Chr(13)
        End If
        If wxParametro580 = "S" And lblDxReferencia1.Text = "" Then
           sMensaje = sMensaje + "Por favor debe ingresar: Dx QUE REFIERE (Ficha 'Cita')" + Chr(13)
        End If
    End If
    If sMensaje <> "" Then
       ValidaDatosObligatoriosReferencias = False
       lcMensaje = lcMensaje & sMensaje
    End If
End Function

Function ValidarReglas() As Boolean
    Dim rsCitas  As Recordset
    Dim lcMensaje As String
    ValidarReglas = False
    If Me.ucEPS1.Visible = True And mi_Opcion = sghModificar Then
        Dim oConexion As New Connection
        oConexion.CommandTimeout = 900
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighEntidades.CadenaConexion
        Set rsCitas = mo_AdminFacturacion.DevuelveSiPagoConsultaMedicaEnCaja(Me.idAtencion, "1", oConexion)
        rsCitas.Filter = "idTipoFinanciamiento=1"
        If rsCitas.RecordCount > 0 Then
           If rsCitas!idestadofacturacion = 4 Then
              MsgBox "Ya pagó en CAJA" & Chr(13) & "debe ANULAR Boleta antes de usar MODIFICACION DE CITA", vbInformation, ""
              Exit Function
           End If
        End If
        rsCitas.Close
        oConexion.Close
        Set oConexion = Nothing
    End If
    If wxParametro540 <> "S" Then
        If mo_Pacientes.idTipoNumeracion > 3 Then
           MsgBox "Solo puede usar HISTORIA FINAL", vbInformation, Me.Caption
           Exit Function
        End If
    End If
    'debb-25/08/2016
    If Len(mo_lcLlegoAlMaximoCuposSIS) > 10 And (mi_Opcion = sghAgregar Or mi_Opcion = sghModificar) _
       And Val(mo_cmbIdFuentesFinanciamiento.BoundText) = sghFuenteFinanciamiento.sghFFSIS Then
       MsgBox mo_lcLlegoAlMaximoCuposSIS, vbInformation, Me.Caption
       Exit Function
    End If
   
    If Not Me.ucPacientesDetalle1.ValidarReglas(mo_Pacientes) Then
        Me.tabAdmision.Tab = 0
        Exit Function
    End If
    '
    If Not mo_AdminAdmision.ValidaEdadMaximaYSexoSegunServicioHosp(Val(txtEdadEnDias.Text), _
                  Val(mo_cmbIdTipoEdad.BoundText), mo_Pacientes.idTipoSexo, Val(mo_cmbIdServicio.BoundText), True) Then
        Me.tabAdmision.Tab = 0
        Exit Function
    End If
    'kike 2017
    If mo_Pacientes.IdDistritoDomicilio = 0 And mo_Pacientes.IdPaisDomicilio = 166 Then
      MsgBox "Por favor elija el DISTRITO DEL DOMICILIO(Ficha 1.1)", vbInformation, Me.Caption
      Me.tabAdmision.Tab = 0
      Me.ucPacientesDetalle1.SetPestaniaTabPaciente 0
      On Error Resume Next
      Me.ucPacientesDetalle1.SetFocusOnDepartamentoDomicilio
      SendKeys "{tab}"
      Exit Function
    End If
    
    If mi_Opcion = sghAgregar Then
         
        '***Verifica si tiene una Cita en el mismo Servicio
        '***verifica que otro Paciente ocupe la HORA
        If chkPacienteNuevo.Value = 1 Then
           mo_Pacientes.idPaciente = 0
           ml_IdPaciente = 0
        Else
             lcMensaje = mo_AdminFacturacion.DevuelveSiElPacienteFallecioOhistoriaPasoPasivo(mo_Pacientes.idPaciente)
             If lcMensaje <> "" Then
               MsgBox lcMensaje, vbInformation, Me.Caption
               Exit Function
             End If
        End If
        Set rsCitas = mo_AdminAdmision.PacienteTieneCita(Me.txtFechaIngreso, Val(mo_cmbIdServicio.BoundText), 0)
        If rsCitas.RecordCount > 0 Then
           'Verifica si tiene una Cita en el mismo Servicio
           If Me.idPaciente > 0 Then
                rsCitas.MoveFirst
                rsCitas.Find "idPaciente=" & Me.idPaciente
                If Not rsCitas.EOF Then
                    MsgBox "El paciente ya tiene una cita en este servicio para la fecha indicada (" & rsCitas!HoraInicio & " - " & rsCitas!HoraFin & ")", vbInformation, Me.Caption
                    Exit Function
                End If
           End If
           'verifica que otro Paciente ya ocupe la HORA de la CITA
           If mi_Opcion = sghAgregar Then
                rsCitas.MoveFirst
                rsCitas.Find "horaInicio='" & Me.txtHoraInicio.Text & "'"
                If Not rsCitas.EOF Then
                     'debb-13/05/2016 (inicio)
                     'MsgBox "Ya existe CITA para otro paciente en esa Hora: " & Me.txtHoraInicio.Text, vbInformation, Me.Caption
                     'If mo_lbEsCitaAdicional = True Then
                        If MsgBox("Ya existe CITA ADICIONAL para otro paciente en esa Hora: " & Me.txtHoraInicio.Text & Chr(13) & "Continua con la siguiente HORA ?", vbQuestion + vbYesNo, "") = vbYes Then
                            Dim lcNuevaHoraFinal As String, lcTiempoAtencion As String
                            Dim daHoraFin  As Date
                            lcTiempoAtencion = DateDiff("n", txtHoraInicio.Text, Me.txtHoraFin.Text)
                            daHoraFin = DateAdd("n", lcTiempoAtencion, CDate(Me.txtHoraFin.Text))
                            txtHoraInicio.Text = txtHoraFin.Text
                            Me.txtHoraFin.Text = Format(daHoraFin, sighEntidades.DevuelveHoraSoloFormato_HM)
                        End If
                     'Else
                     '   MsgBox "Ya existe CITA para otro paciente en esa Hora: " & Me.txtHoraInicio.Text, vbInformation, Me.Caption
                     'End If
                     'debb-13/05/2016 (fin)
                     Exit Function
                End If
           End If
        End If
        rsCitas.Close
        
        
        
        'Verifica si tien Cita en cualquier Servicio, en el mismo dia
        If Me.idPaciente > 0 Then
            Set rsCitas = mo_AdminAdmision.PacienteTieneCita(Me.txtFechaIngreso, 0, Me.idPaciente)
            If Not (rsCitas.EOF And rsCitas.BOF) Then
                lcMensaje = "El paciente ya tiene una cita en el servicio: " & rsCitas.Fields!codServicio & " - " & rsCitas.Fields!DServicio & Chr(13) & " para la fecha indicada (" & rsCitas!HoraInicio & " - " & rsCitas!HoraFin & ")"
                If lbUsuarioAutorizadoAregistrarCitasRepetidas = True Then
                    If rsCitas!HoraInicio = txtHoraInicio.Text Then
                        MsgBox lcMensaje + Chr(13) + Chr(13) + "Puede hacer la CITA pero en otra hora", vbInformation, Me.Caption
                        Exit Function
                    Else
                        If MsgBox(lcMensaje & Chr(13) & "¿Desea AGREGAR LA CITA?", vbQuestion + vbYesNo, "") = vbNo Then
                           Exit Function
                        End If
                    End If
                Else
                   MsgBox lcMensaje, vbInformation, Me.Caption
                   Exit Function
                End If
            End If
            rsCitas.Close
        End If
        
    End If
    
    
    'WCG_2006
    If (mi_Opcion = sghAgregar) Or (mi_Opcion = sghModificar) Then
        If Val(mo_cmbIdTipoConsulta.BoundText) = 0 Then
            MsgBox "Por favor ingrese el tipo de consulta", vbInformation, Me.Caption
            Exit Function
        End If
    End If
    'WCG_2006
   
    If Me.ucPacientesDetalle1.FechaNacimiento <> sighEntidades.FECHA_VACIA_DMY Then
        If CDate(Me.ucPacientesDetalle1.FechaNacimiento) > CDate(Me.txtFechaIngreso) Then
            MsgBox "La fecha de ingreso no puede ser menor que la fecha de nacimiento", vbExclamation, Me.Caption
            Exit Function
        End If
    End If
    If wxParametro302 = "S" And Val(mo_cmbIdFuentesFinanciamiento.BoundText) = sghFuenteFinanciamiento.sghFFSIS And _
            mi_Opcion = sghEliminar Then
            Set rsCitas = mo_ReglasSISgalenhos.SisFuaAtencionSeleccionarPorCuenta(Val(Me.txtNroCuenta.Text))
            If rsCitas.RecordCount > 0 Then
               MsgBox "El formato FUA ya fué generado: " & rsCitas.Fields!fuaDisa & "-" & rsCitas!fuaLote & "-" & _
                      rsCitas!FuaNumero & Chr(13) & "Debe eliminar el formato FUA (módulo: SIS, opción: Formato FUA)", _
                      vbInformation, Me.Caption
               
               Exit Function
            End If
    End If
   
    If mi_Opcion = sghAgregar Then 'Frank 2508
        If wxParametro302 = "S" And Val(mo_cmbIdFuentesFinanciamiento.BoundText) = sghFuenteFinanciamiento.sghFFSIS Then
           lcMensaje = mo_ReglasSISgalenhos.ChequeaCodigoEstablecimientoAdscripcion(lcCodigoEstablecimientoAdscripcionSIS, _
                                            sghConsultaExterna, _
                                            mo_AdminAdmision.TiposOrigenAtencionDevuelveIdSis(Val(mo_cmbIdViasAdmision.BoundText)), _
                                            "")
           If lcMensaje <> "" Then
              Me.tabAdmision.Tab = 1
              MsgBox lcMensaje, vbInformation, Me.Caption
              
              CargarAutomaticamenteEstablecimientoReferenciaSIS ''''' Frank 2608

              Exit Function
           End If
        End If
    End If
    
    If mi_Opcion = sghEliminar Or mi_Opcion = sghModificar Then
        Dim oAtencionesCE As New AtencionesCE
        Dim oDOAtencionesCE As New DOAtencionesCE
        Dim oConexionExterna As New Connection
        oConexionExterna.CommandTimeout = 300
        oConexionExterna.CursorLocation = adUseClient
        oConexionExterna.Open wxParametroJAMO
        oDOAtencionesCE.idAtencion = Me.idAtencion   'ml_idAtencion
        Set oAtencionesCE.Conexion = oConexionExterna
        If oAtencionesCE.SeleccionarPorId(oDOAtencionesCE) = True Then
            If Len(Trim(oDOAtencionesCE.CitaDiagMed)) > 0 Then
                'MsgBox "Ya se registró la Atención", vbInformation, Me.Caption
                If mi_Opcion = sghEliminar Then
                    MsgBox "No puede Eliminar la Cita, la Atención ya fue registrada. Revise el registro de atenciones", vbExclamation, Me.Caption
                End If
                If mi_Opcion = sghModificar Then
                    MsgBox "No puede Modificar la Cita, la Atención ya fue registrada. Revise el registro de atenciones", vbExclamation, Me.Caption
                End If
                Exit Function
            Else
                If Not IsNull(oDOAtencionesCE.TriajeFecha) Then
                    If mi_Opcion = sghEliminar Then
                        MsgBox "No puede Eliminar la Cita, el paciente ya paso por Triaje.", vbExclamation, Me.Caption
                    End If
                    If mi_Opcion = sghModificar Then
                        MsgBox "No puede Modificar la Cita, el paciente ya paso por Triaje.", vbExclamation, Me.Caption
                    End If
                    Exit Function
                End If
            End If
        End If
        Set oDOAtencionesCE = Nothing
        Set oAtencionesCE = Nothing
        oConexionExterna.Close
        Set oConexionExterna = Nothing
    End If
    
    ValidarReglas = True
    Set rsCitas = Nothing
End Function

Public Sub CargarAutomaticamenteEstablecimientoReferenciaSIS() 'Frank 2808
    If lcBuscaParametro.SeleccionaFilaParametro(326) = "S" And lcCodigoEstablecimientoAdscripcionSIS <> "" Then
       Dim lcCodigoSis As String
       Dim lcEstablecimientoOrigen As String
       Dim DOEstablecimiento As New DOEstablecimiento
       Dim oRsEstabNoMINSA As Recordset
       Dim lnIdOrigenDelPacienteDesdeFUA As Long
       lnIdOrigenDelPacienteDesdeFUA = mo_AdminAdmision.TiposOrigenAtencionDevuelveIdSis(Val(mo_cmbIdViasAdmision.BoundText))
       
       If Val(lcBuscaParametro.SeleccionaFilaParametro(280)) <> Val(lcCodigoEstablecimientoAdscripcionSIS) Then
          If lcBuscaParametro.SeleccionaFilaParametro(282) <> "S" Then 'Hospital
               If Not (lnIdOrigenDelPacienteDesdeFUA = "4" Or lnIdOrigenDelPacienteDesdeFUA = "6") Then 'Referido CE, ContraReferido
                    mo_cmbIdViasAdmision.BoundText = "12"
                    If mo_AdminServiciosComunes.EstablecimientosSeleccionarPorCodigo(Right(lcCodigoEstablecimientoAdscripcionSIS, 5), DOEstablecimiento) = True Then
                        mo_cmbIdTipoReferenciaOrigen.BoundText = 1 'MINSA
                        txtIdEstablecimientoOrigen.Text = DOEstablecimiento.Codigo
                        txtIdEstablecimientoOrigen.Tag = DOEstablecimiento.IdEstablecimiento
                        txtNombreOrigenReferencia.Text = DOEstablecimiento.nombre
                    Else
                        Set oRsEstabNoMINSA = mo_AdminServiciosComunes.EstablecimientosNoMinsaSeleccionarPorCodigo(Right(lcCodigoEstablecimientoAdscripcionSIS, 5))
                        If oRsEstabNoMINSA.RecordCount > 0 Then
                            oRsEstabNoMINSA.MoveFirst
                            mo_cmbIdTipoReferenciaOrigen.BoundText = 2 'NO MINSA
                            txtIdEstablecimientoOrigen.Text = oRsEstabNoMINSA.Fields!Codigo
                            txtIdEstablecimientoOrigen.Tag = oRsEstabNoMINSA.Fields!IdEstablecimientoNoMINSA
                            txtNombreOrigenReferencia.Text = oRsEstabNoMINSA.Fields!nombre
                        End If
                        Set oRsEstabNoMINSA = Nothing
                    End If
               End If
          End If
       End If
       Set DOEstablecimiento = Nothing
    End If
End Sub

Function ConvertirAMinutos(sHora As String) As Long
Dim sHoras() As String
        
        sHoras = Split(sHora, ":")
        ConvertirAMinutos = Val(sHoras(0)) * 60 + Val(sHoras(1))
        
End Function

Sub CargaDatosAlObjetosDeDatos()
    'Limpia Dx
    Set mo_Diagnosticos = Nothing
    '
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LA CUENTA ATENCION
    '---------------------------------------------------------------------------------
   With mo_CuentasAtencion
           Select Case mi_Opcion
           Case sghAgregar, sghModificar, sghEliminar
                .idCuentaAtencion = Me.idCuentaAtencion
                .idPaciente = Me.idPaciente
                .TotalAsegurado = 0
                .TotalExonerado = 0
                .TotalPagado = 0
                .TotalPorPagar = 0
                'WCG 10/06
                .FechaApertura = Me.txtFechaIngreso.Text
                .HoraApertura = Me.txtHoraInicio.Text
                .fechaCierre = 0
                .HoraCierre = ""
                .IdUsuarioAuditoria = ml_idUsuario
                .idEstado = sghEstadoCuenta.sghAbierto
           End Select
   End With
   
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LA ATENCION
    '---------------------------------------------------------------------------------
   With mo_Atenciones
   
            .idAtencion = Me.idAtencion
            .IdEspecialidadMedico = mo_cmbIdEspecialidadMedico.BoundText
            .IdMedicoIngreso = Me.idMedico
            .IdServicioIngreso = Val(mo_cmbIdServicio.BoundText)
            .IdOrigenAtencion = Val(mo_cmbIdViasAdmision.BoundText)
            
           
           
           .HoraIngreso = Me.txtHoraInicio.Text
           .FechaIngreso = Me.txtFechaIngreso.Text
           .idTipoServicio = mo_cmbIdTipoServicio.BoundText
           .Edad = Me.txtEdadEnDias.Text
           .idTipoEdad = mo_cmbIdTipoEdad.BoundText
           .idPaciente = Me.idPaciente
            .IdMedicoEgreso = 0
            .HoraEgreso = ""
            .fechaEgreso = 0
               
            If Me.chkPacienteNuevo = 1 Then
                .IdTipoCondicionALEstab = 1
                .IdTipoCondicionAlServicio = 1
            Else
                If mi_Opcion = sghAgregar Then
                    Dim lnIdCondicionAlServicio As Long, lnIdCondicionAlEstablecimiento As Long
                    mo_AdminServiciosComunes.TiposCondicionPacienteCondicionAlEstablecimientoYservicio lnIdCondicionAlEstablecimiento, lnIdCondicionAlServicio, Me.idPaciente, Format(Me.txtFechaIngreso, sighEntidades.DevuelveFechaSoloFormato_DMY), Me.idAtencion, Val(mo_cmbIdServicio.BoundText)
                    .IdTipoCondicionALEstab = lnIdCondicionAlEstablecimiento
                    .IdTipoCondicionAlServicio = lnIdCondicionAlServicio
                End If
            End If
            .IdTipoGravedad = 0
            .IdUsuarioAuditoria = ml_idUsuario
            .IdFormaPago = IIf(mo_cmbIdFormaPago.BoundText = "", 0, Val(mo_cmbIdFormaPago.BoundText))
            .IdFuenteFinanciamiento = IIf(mo_cmbIdFuentesFinanciamiento.BoundText = "", 0, Val(mo_cmbIdFuentesFinanciamiento.BoundText))
            idFormaPagoProvisional = IIf(mo_cmbIdFormaPago.BoundText = "", 0, Val(mo_cmbIdFormaPago.BoundText))
            .IdEstadoAtencion = sghEstadoTabla.sghRegistrado
            If ucEPS1.Visible = True Then
               .EpsPorcentaje = ucEPS1.Porcentaje
            Else
               .EpsPorcentaje = 0
            End If
            
   End With

    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LA CITA
    '---------------------------------------------------------------------------------
   With mo_Cita
           .IdCita = Me.IdCita
           .fecha = Me.txtFechaIngreso.Text
           .HoraInicio = Me.txtHoraInicio.Text
           .HoraFin = Me.txtHoraFin.Text
           .idMedico = Me.idMedico
           .IdEspecialidad = mo_cmbIdEspecialidadMedico.BoundText
           .idPaciente = Me.idPaciente
           .IdServicio = mo_cmbIdServicio.BoundText
           .IdEstadoCita = 1    'CITA SEPARADA
           .idAtencion = Me.idAtencion
           .IdProgramacion = Me.IdProgramacion
           .idProducto = Val(mo_cmbIdTipoConsulta.BoundText)
           .IdUsuarioAuditoria = ml_idUsuario
           If mi_Opcion = sghAgregar Then
                .FechaSolicitud = Format(Now, sighEntidades.DevuelveFechaSoloFormato_DMY)
                .HoraSolicitud = Format(Now, sighEntidades.DevuelveHoraSoloFormato_HM)
           End If
           .EsCitaAdicional = mo_lbEsCitaAdicional
   End With


    '---------------------------------------------------------------------------------
    '           CARGA DATOS DEL PACIENTE
    '---------------------------------------------------------------------------------
    Me.ucPacientesDetalle1.idUsuario = ml_idUsuario
    Me.ucPacientesDetalle1.CargarDatosAlObjetoDatos mo_Pacientes, mo_Historia, mo_DoPacientesDatosAdd
    
    '---------------------------------------------------------------------------------
    '           COMPLETA LOS DATOS DE LA ATENCION
    '---------------------------------------------------------------------------------
    With mo_DoAtencionDatosAdicionales
        .IdTipoReferenciaOrigen = Val(mo_cmbIdTipoReferenciaOrigen.BoundText)
        If .IdTipoReferenciaOrigen = 1 Then
            .IdEstablecimientoOrigen = Val(Me.txtIdEstablecimientoOrigen.Tag)
            .IdEstablecimientoNoMinsaOrigen = 0
        Else
            .IdEstablecimientoOrigen = 0
            .IdEstablecimientoNoMinsaOrigen = Val(Me.txtIdEstablecimientoOrigen.Tag)
        End If
        .NroReferenciaOrigen = txtReferenciaO.Text
        .DireccionDomicilio = mo_Pacientes.DireccionDomicilio
        If mi_Opcion = sghAgregar Or mi_Opcion = sghModificar Then
           If Val(mo_cmbIdFuentesFinanciamiento.BoundText) = sghFuenteFinanciamiento.sghFFSIS Then
                If lnAfiliacionSIS4 = 0 Then
                   mo_ReglasSISgalenhos.SisFiliacionesDevuelveKEY lnAfiliacionSIS4, lcSIScodigo, _
                                        mo_Pacientes.ApellidoPaterno, mo_Pacientes.ApellidoMaterno, _
                                        mo_Pacientes.PrimerNombre, mo_Pacientes.FechaNacimiento, _
                                        lcCodigoEstablecimientoAdscripcionSIS
                End If
                .idSiaSis = lnAfiliacionSIS4
                .SisCodigo = lcSIScodigo
                If mi_Opcion = sghAgregar Then
                   .FuaCodigoPrestacion = ""
                   'SCCQ 23-04-2021 Cambio 64 Inicio (Reversion Cambio 46)
                                
                    '.FuaCodigoPrestacion1 = Me.ucSISfuaCodPrestacion1.CodigoPrestacion 'HRA 10/12/2020 Cambio 46
                    'SCCQ 23-04-2021 Cambio 64 Fin (Reversion Cambio 46)
                ElseIf wxParametro302 = "S" And Me.ucSISfuaCodPrestacion1.CodigoPrestacion <> "" Then
                   .FuaCodigoPrestacion = Me.ucSISfuaCodPrestacion1.CodigoPrestacion
                   'SCCQ 23-04-2021 Cambio 64 Inicio (Reversion Cambio 46)
                  '.FuaCodigoPrestacion1 = Me.ucSISfuaCodPrestacion1.CodigoPrestacion 'HRA 10/12/2020 Cambio 46
                   'SCCQ 23-04-2021 Cambio 64 Fin (Reversion Cambio 46)
                End If
                
           Else
                .idSiaSis = 0
                .FuaCodigoPrestacion = ""
                .SisCodigo = ""
                .sisAfiliacion = ""
           End If
           '
           .referenciaOservicio = PVcomboBoxDevuelveEleccion(cmbServicioReferenciaO)
           .referenciaOidDiagnostico = Val(txtDxReferencia.Tag)
           'FRANKLIN 2017
           .ReferenciaMedicoOColeg = Me.txtMedicoRef.Text
           If Trim(Me.cmbMedicoRef.Text) <> "" Then
              .ReferenciaMedicoOIdcolegio = Trim(Split(Me.cmbMedicoRef.Text, " = ")(0))
           Else
              .ReferenciaMedicoOIdcolegio = ""
           End If
           '
           
        End If
    End With
    
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE DIAGNOSTICOS DE INGRESO
    '---------------------------------------------------------------------------------
    
    '---------------------------------------------------------------------------------
    '           CARGA SERVICIOS
    '---------------------------------------------------------------------------------
    Select Case mi_Opcion
    Case sghAgregar
        CargaDeServiciosAFacturar
    Case sghModificar
        CargaDeServiciosAFacturar
    Case sghEliminar
    End Select
    '
    mo_Atenciones.IdFormaPago = idFormaPagoProvisional
    mo_DoAtencionDatosAdicionales.RecienNacido = sighEntidades.CalculaSiEsRecienNacido(mo_Pacientes.FechaNacimiento, CDate(mo_Atenciones.FechaIngreso & " " & mo_Atenciones.HoraIngreso))
    '
    Me.UcPacientesSunasa1.idUsuario = ml_idUsuario
    Me.UcPacientesSunasa1.CargarDatosAlObjetoDatos oDoSunasaPacientesHistoricos
    '
End Sub



Sub CargaDeServicioModificados()
        
        Set mo_FacturacionServiciosPorEliminar = New Collection
        Set mo_FacturacionServicios = New Collection
        
        '**********Si son diferentes significa que se ha modificado
        '**********Si se cambio de FORMA DE PAGO y EL Nº DE ORDEN AUN NO SE HA PAGADO
        If Val(Me.cmbIdTipoConsulta.Tag) <> Val(mo_cmbIdTipoConsulta.BoundText) Or lnFormaPagoAnterior <> Val(mo_cmbIdFormaPago.BoundText) Then
            Dim oConsulta As New DOFacturacionServicios
            oConsulta.idAtencion = Me.idAtencion
            oConsulta.idProducto = Val(Me.cmbIdTipoConsulta.Tag) 'Para eliminar el tipo de consulta anterior
            If oConsulta.idProducto <> 0 And (oConsulta.idestadofacturacion = 0 Or oConsulta.idestadofacturacion = 1) Then    'Si es que existia un producto
                mo_FacturacionServiciosPorEliminar.Add oConsulta
            End If
            mo_FacturacionServicios.Add CargarTipoDeConsulta()
        End If
        
        'Si son diferentes significa que se ha modificado
        If Val(Me.chkNuevoCarne.Tag) <> Val(Me.chkNuevoCarne.Value) Or lnFormaPagoAnterior <> Val(mo_cmbIdFormaPago.BoundText) Then
            Dim oNuevoCarne As New DOFacturacionServicios
            oNuevoCarne.idAtencion = Me.idAtencion
            oNuevoCarne.idProducto = mo_AdminFacturacion.ObtenerCodigoDeNuevoCarnet()
            If oNuevoCarne.idProducto <> 0 And (oConsulta.idestadofacturacion = 0 Or oConsulta.idestadofacturacion = 1) Then    'Si es que existia un producto
                mo_FacturacionServiciosPorEliminar.Add oNuevoCarne
            End If
            If Me.chkNuevoCarne.Value = 0 Then  'Si ahora esta desmarcado hay que eliminar el anterior
            Else
                mo_FacturacionServicios.Add CargarNuevoCarne()
            End If
        End If
        
        'Si son diferentes significa que se ha modificado
        If Val(Me.chkDuplicadoCarne.Tag) <> Val(Me.chkDuplicadoCarne.Value) Or lnFormaPagoAnterior <> Val(mo_cmbIdFormaPago.BoundText) Then
            Dim oDuplicadoCarne As New DOFacturacionServicios
            oDuplicadoCarne.idAtencion = Me.idAtencion
            oDuplicadoCarne.idProducto = mo_AdminFacturacion.ObtenerCodigoDeDuplicadoCarnet()
            If oDuplicadoCarne.idProducto <> 0 And (oConsulta.idestadofacturacion = 0 Or oConsulta.idestadofacturacion = 1) Then    'Si es que existia un producto
                mo_FacturacionServiciosPorEliminar.Add oDuplicadoCarne
            End If
            If Me.chkDuplicadoCarne.Value = 0 Then  'Si ahora esta desmarcado hay que eliminar el anterior
            Else
                mo_FacturacionServicios.Add CargarDuplicadoCarne()
            End If
        End If
        
        'Si son diferentes significa que se ha modificado
        If Val(Me.chkNuevoFolder.Tag) <> Val(Me.chkNuevoFolder.Value) Or lnFormaPagoAnterior <> Val(mo_cmbIdFormaPago.BoundText) Then
            Dim oFolder As New DOFacturacionServicios
            oFolder.idAtencion = Me.idAtencion
            oFolder.idProducto = mo_AdminFacturacion.ObtenerCodigoDeFolder()
            If oFolder.idProducto <> 0 And (oConsulta.idestadofacturacion = 0 Or oConsulta.idestadofacturacion = 1) Then    'Si es que existia un producto
                mo_FacturacionServiciosPorEliminar.Add oFolder
            End If
            If Me.chkNuevoFolder.Value = 0 Then  'Si ahora esta desmarcado hay que eliminar el anterior
            Else
                mo_FacturacionServicios.Add CargarNuevoFolder()
            End If
        End If
        
        
End Sub
Function CargarTipoDeConsulta() As DOFacturacionServicios
Dim oConsulta As New DOFacturacionServicios
Dim oRsBuscaSeguro As New ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim PrSeguro As Double
        With oConsulta
            .idAtencion = Me.idAtencion
            .IdFacturacionServicio = 0  'Id
            .idTipoFinanciamiento = Val(mo_cmbIdFormaPago.BoundText)
            .IdFuenteFinanciamiento = Val(mo_cmbIdFuentesFinanciamiento.BoundText)
            .Cantidad = 1
            .idProducto = Val(mo_cmbIdTipoConsulta.BoundText)
            .IdUsuarioAuditoria = ml_idUsuario
            .idestadofacturacion = sghEstadoFacturacion.sghRegistrado
            .FechaAutorizaPendiente = 0
            .FechaAutorizaSeguro = 0
            .IdCentroCosto = 0
            .IdUsuarioAutorizaPendiente = 0
            .IdUsuarioAutorizaSeguro = 0
            .PrecioUnitario = 0
            .TotalPorPagar = 0
            .idPuntoCarga = 6
            '********pone Seguros en forma automatica, sin necesidad de ir a SEGUROS-inicio
            If mi_Opcion = sghModificar Then
               .IdOrden = lnIdFactServicios
            End If
            PrSeguro = 0
            oConexion.Open sighEntidades.CadenaConexion
            oConexion.CursorLocation = adUseClient
            If Val(mo_cmbIdFormaPago.BoundText) = 9 Then
               'Si es EXONERACIONES tomará el PRECIO de un Paciente Normal
               Set oRsBuscaSeguro = mo_AdminFacturacion.CatalogoServiciosHospSeleccionarXidProductoIdTipoFinanciamiento(.idProducto, 1, oConexion)
            Else
               Set oRsBuscaSeguro = mo_AdminFacturacion.CatalogoServiciosHospSeleccionarXidProductoIdTipoFinanciamiento(.idProducto, Val(mo_cmbIdFormaPago.BoundText), oConexion)
            End If
            If oRsBuscaSeguro.RecordCount > 0 Then
               PrSeguro = oRsBuscaSeguro.Fields!PrecioUnitario
            End If
            oRsBuscaSeguro.Close
            Set oRsBuscaSeguro = Nothing
            oConexion.Close
            Set oConexion = Nothing
            .idTipoFinanciamiento = Val(mo_cmbIdFormaPago.BoundText)
            .CantidadSIS = 0
            .precioSIS = 0
            .ImporteSIS = 0
            .CantidadSOAT = 0
            .PrecioSOAT = 0
            .ImporteSOAT = 0
            .importeEXO = 0
            .cantidadConv = 0
            .precConv = 0
            .ImporteConv = 0
            Select Case Val(mo_cmbIdFormaPago.BoundText)
            Case 1  'Contado
            Case 2  'SIS
                 If PrSeguro > 0 Then
                    .CantidadSIS = 1
                    .precioSIS = PrSeguro
                    .ImporteSIS = PrSeguro
                    .idestadofacturacion = 10
                    .FechaAutorizaSeguro = Now
                 Else
                 End If
            Case 3  'Soat
                 If PrSeguro > 0 Then
                    .CantidadSOAT = 1
                    .PrecioSOAT = PrSeguro
                    .ImporteSOAT = PrSeguro
                    .idestadofacturacion = 10
                    .FechaAutorizaSeguro = Now
                 Else
                    .idTipoFinanciamiento = 1
                    'mo_Atenciones.IdFormaPago = 1
                    'cmbFormaPago.BoundText = "1"
                    'mo_cmbIdFormaPago.BoundText = "1"
                 End If
            Case 4  'Convenios
                 If PrSeguro > 0 Then
                    .cantidadConv = 1
                    .precConv = PrSeguro
                    .ImporteConv = PrSeguro
                    .idestadofacturacion = 10
                    .PrecioUnitario = PrSeguro
                    .FechaAutorizaSeguro = Now
                 Else
                    .idTipoFinanciamiento = 1
                    'mo_Atenciones.IdFormaPago = 1
                    'cmbFormaPago.BoundText = "1"
                    'mo_cmbIdFormaPago.BoundText = "1"
                 End If
            Case 9  'Exonerados
                .importeEXO = PrSeguro
                .idestadofacturacion = 10
                .idTipoFinanciamiento = 1
                .FechaAutorizaEXO = Now
            End Select
            '********pone Seguros en forma automatica, sin necesidad de ir a SEGUROS-fin
        End With

        Set CargarTipoDeConsulta = oConsulta


End Function
Function CargarNuevoCarne() As DOFacturacionServicios
Dim oNuevoCarne As New DOFacturacionServicios
Dim oRsBuscaSeguro As New ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim PrSeguro As Double
            
            With oNuevoCarne
                .idAtencion = Me.idAtencion
                .IdFacturacionServicio = 0 'Id
                .IdFuenteFinanciamiento = Val(mo_cmbIdFuentesFinanciamiento.BoundText)
                .idTipoFinanciamiento = Val(mo_cmbIdFormaPago.BoundText)
                .Cantidad = 1
                .idProducto = mo_AdminFacturacion.ObtenerCodigoDeNuevoCarnet()
                .IdUsuarioAuditoria = ml_idUsuario
                .idestadofacturacion = sghEstadoFacturacion.sghRegistrado
                .FechaAutorizaPendiente = 0
                .FechaAutorizaSeguro = 0
                .IdCentroCosto = 0
                .IdUsuarioAutorizaPendiente = 0
                .IdUsuarioAutorizaSeguro = 0
                .PrecioUnitario = 0
                .TotalPorPagar = 0
                .idPuntoCarga = 6
                '********pone Seguros en forma automatica, sin necesidad de ir a SEGUROS-inicio
                If mi_Opcion = sghModificar Then
                   .IdOrden = lnIdFactServicios
                End If
                PrSeguro = 0
                oConexion.Open sighEntidades.CadenaConexion
                oConexion.CursorLocation = adUseClient
                If Val(mo_cmbIdFormaPago.BoundText) = 9 Then
                   'Si es EXONERACIONES tomará el PRECIO de un Paciente Normal
                   Set oRsBuscaSeguro = mo_AdminFacturacion.CatalogoServiciosHospSeleccionarXidProductoIdTipoFinanciamiento(.idProducto, 1, oConexion)
                Else
                   Set oRsBuscaSeguro = mo_AdminFacturacion.CatalogoServiciosHospSeleccionarXidProductoIdTipoFinanciamiento(.idProducto, Val(mo_cmbIdFormaPago.BoundText), oConexion)
                End If
                If oRsBuscaSeguro.RecordCount > 0 Then
                   PrSeguro = oRsBuscaSeguro.Fields!PrecioUnitario
                End If
                oRsBuscaSeguro.Close
                Set oRsBuscaSeguro = Nothing
                oConexion.Close
                Set oConexion = Nothing
                .idTipoFinanciamiento = Val(mo_cmbIdFormaPago.BoundText)
                .CantidadSIS = 0
                .precioSIS = 0
                .ImporteSIS = 0
                .CantidadSOAT = 0
                .PrecioSOAT = 0
                .ImporteSOAT = 0
                .importeEXO = 0
                .cantidadConv = 0
                .precConv = 0
                .ImporteConv = 0
                Select Case Val(mo_cmbIdFormaPago.BoundText)
                Case 1  'Contado
                Case 2  'SIS
                     If PrSeguro > 0 Then
                        .CantidadSIS = 1
                        .precioSIS = PrSeguro
                        .ImporteSIS = PrSeguro
                      '  .cantidad = 0
                        .idestadofacturacion = 10
                        .FechaAutorizaSeguro = Now
                     Else
                        .idTipoFinanciamiento = 1
                        'mo_Atenciones.IdFormaPago = 1
                        'cmbFormaPago.BoundText = "1"
                        'mo_cmbIdFormaPago.BoundText = "1"
                     End If
                Case 3  'Soat
                     If PrSeguro > 0 Then
                        .CantidadSOAT = 1
                        .PrecioSOAT = PrSeguro
                        .ImporteSOAT = PrSeguro
                       ' .cantidad = 0
                        .idestadofacturacion = 10
                        .FechaAutorizaSeguro = Now
                     Else
                        .idTipoFinanciamiento = 1
                        'mo_Atenciones.IdFormaPago = 1
                        'cmbFormaPago.BoundText = "1"
                        'mo_cmbIdFormaPago.BoundText = "1"
                     End If
                Case 4  'Convenios
                     If PrSeguro > 0 Then
                        .cantidadConv = 1
                        .precConv = PrSeguro
                        .ImporteConv = PrSeguro
                        .idestadofacturacion = 10
                        .FechaAutorizaSeguro = Now
                     Else
                        .idTipoFinanciamiento = 1
                        'mo_Atenciones.IdFormaPago = 1
                        'mo_cmbIdFormaPago.BoundText = "1"
                     End If
                Case 9  'Exonerados
                    .importeEXO = PrSeguro
                    .idestadofacturacion = 10
                    .idTipoFinanciamiento = 1
                   ' .cantidad = 0
                    .FechaAutorizaEXO = Now
                End Select
                '********pone Seguros en forma automatica, sin necesidad de ir a SEGUROS-fin
            End With
            
            Set CargarNuevoCarne = oNuevoCarne

End Function
Function CargarDuplicadoCarne() As DOFacturacionServicios
Dim oDuplicadoCarne As New DOFacturacionServicios
Dim oRsBuscaSeguro As New ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim PrSeguro As Double
            
            With oDuplicadoCarne
                .idAtencion = Me.idAtencion
                .IdFacturacionServicio = 0
                .IdFuenteFinanciamiento = Val(mo_cmbIdFuentesFinanciamiento.BoundText)
                .idTipoFinanciamiento = Val(mo_cmbIdFormaPago.BoundText)
                .Cantidad = 1
                .idProducto = mo_AdminFacturacion.ObtenerCodigoDeDuplicadoCarnet()
                .IdUsuarioAuditoria = ml_idUsuario
                .idestadofacturacion = sghEstadoFacturacion.sghRegistrado
                .FechaAutorizaPendiente = 0
                .FechaAutorizaSeguro = 0
                .IdCentroCosto = 0
                .IdUsuarioAutorizaPendiente = 0
                .IdUsuarioAutorizaSeguro = 0
                .PrecioUnitario = 0
                .TotalPorPagar = 0
                .idPuntoCarga = 6
                '********pone Seguros en forma automatica, sin necesidad de ir a SEGUROS-inicio
                If mi_Opcion = sghModificar Then
                   .IdOrden = lnIdFactServicios
                End If
                PrSeguro = 0
                oConexion.Open sighEntidades.CadenaConexion
                oConexion.CursorLocation = adUseClient
                If Val(mo_cmbIdFormaPago.BoundText) = 9 Then
                   'Si es EXONERACIONES tomará el PRECIO de un Paciente Normal
                   Set oRsBuscaSeguro = mo_AdminFacturacion.CatalogoServiciosHospSeleccionarXidProductoIdTipoFinanciamiento(.idProducto, 1, oConexion)
                Else
                   Set oRsBuscaSeguro = mo_AdminFacturacion.CatalogoServiciosHospSeleccionarXidProductoIdTipoFinanciamiento(.idProducto, Val(mo_cmbIdFormaPago.BoundText), oConexion)
                End If
                If oRsBuscaSeguro.RecordCount > 0 Then
                   PrSeguro = oRsBuscaSeguro.Fields!PrecioUnitario
                End If
                oRsBuscaSeguro.Close
                Set oRsBuscaSeguro = Nothing
                oConexion.Close
                Set oConexion = Nothing
                .idTipoFinanciamiento = Val(mo_cmbIdFormaPago.BoundText)
                .CantidadSIS = 0
                .precioSIS = 0
                .ImporteSIS = 0
                .CantidadSOAT = 0
                .PrecioSOAT = 0
                .ImporteSOAT = 0
                .importeEXO = 0
                .cantidadConv = 0
                .precConv = 0
                .ImporteConv = 0
                Select Case Val(mo_cmbIdFormaPago.BoundText)
                Case 1  'Contado
                Case 2  'SIS
                     If PrSeguro > 0 Then
                        .CantidadSIS = 1
                        .precioSIS = PrSeguro
                        .ImporteSIS = PrSeguro
                      '  .cantidad = 0
                        .idestadofacturacion = 10
                        .FechaAutorizaSeguro = Now
                     Else
                        .idTipoFinanciamiento = 1
                       ' mo_Atenciones.IdFormaPago = 1
                       ' mo_cmbIdFormaPago.BoundText = "1"
                     End If
                Case 3  'Soat
                     If PrSeguro > 0 Then
                        .CantidadSOAT = 1
                        .PrecioSOAT = PrSeguro
                        .ImporteSOAT = PrSeguro
                        .idestadofacturacion = 10
                        .FechaAutorizaSeguro = Now
                     Else
                        .idTipoFinanciamiento = 1
                       ' mo_Atenciones.IdFormaPago = 1
                       ' mo_cmbIdFormaPago.BoundText = "1"
                     End If
                Case 4  'Convenios
                     If PrSeguro > 0 Then
                        .cantidadConv = 1
                        .precConv = PrSeguro
                        .ImporteConv = PrSeguro
                        .idestadofacturacion = 10
                        .FechaAutorizaSeguro = Now
                     Else
                        .idTipoFinanciamiento = 1
                       ' mo_Atenciones.IdFormaPago = 1
                       ' mo_cmbIdFormaPago.BoundText = "1"
                     End If
                Case 9  'Exonerados
                    .importeEXO = PrSeguro
                    .idestadofacturacion = 10
                    .idTipoFinanciamiento = 1
                    .FechaAutorizaEXO = Now
                End Select
                '********pone Seguros en forma automatica, sin necesidad de ir a SEGUROS-fin
                
            End With

            Set CargarDuplicadoCarne = oDuplicadoCarne
                        
End Function
Function CargarNuevoFolder() As DOFacturacionServicios
Dim oFolder As New DOFacturacionServicios
Dim oRsBuscaSeguro As New ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim PrSeguro As Double
            
            With oFolder
                .idAtencion = Me.idAtencion
                .IdFacturacionServicio = 0
                .IdFuenteFinanciamiento = Val(mo_cmbIdFuentesFinanciamiento.BoundText)
                .idTipoFinanciamiento = Val(mo_cmbIdFormaPago.BoundText)
                .Cantidad = 1
                .idProducto = mo_AdminFacturacion.ObtenerCodigoDeFolder()
                .IdUsuarioAuditoria = ml_idUsuario
                .idestadofacturacion = sghEstadoFacturacion.sghRegistrado
                .FechaAutorizaPendiente = 0
                .FechaAutorizaSeguro = 0
                .IdCentroCosto = 0
                .IdUsuarioAutorizaPendiente = 0
                .IdUsuarioAutorizaSeguro = 0
                .PrecioUnitario = 0
                .TotalPorPagar = 0
                .idPuntoCarga = 6
                '********pone Seguros en forma automatica, sin necesidad de ir a SEGUROS-inicio
                If mi_Opcion = sghModificar Then
                   .IdOrden = lnIdFactServicios
                End If
                PrSeguro = 0
                oConexion.Open sighEntidades.CadenaConexion
                oConexion.CursorLocation = adUseClient
                If Val(mo_cmbIdFormaPago.BoundText) = 9 Then
                   'Si es EXONERACIONES tomará el PRECIO de un Paciente Normal
                   Set oRsBuscaSeguro = mo_AdminFacturacion.CatalogoServiciosHospSeleccionarXidProductoIdTipoFinanciamiento(.idProducto, 1, oConexion)
                Else
                   Set oRsBuscaSeguro = mo_AdminFacturacion.CatalogoServiciosHospSeleccionarXidProductoIdTipoFinanciamiento(.idProducto, Val(mo_cmbIdFormaPago.BoundText), oConexion)
                End If
                If oRsBuscaSeguro.RecordCount > 0 Then
                   PrSeguro = oRsBuscaSeguro.Fields!PrecioUnitario
                End If
                oRsBuscaSeguro.Close
                Set oRsBuscaSeguro = Nothing
                oConexion.Close
                Set oConexion = Nothing
                .idTipoFinanciamiento = Val(mo_cmbIdFormaPago.BoundText)
                .CantidadSIS = 0
                .precioSIS = 0
                .ImporteSIS = 0
                .CantidadSOAT = 0
                .PrecioSOAT = 0
                .ImporteSOAT = 0
                .importeEXO = 0
                .cantidadConv = 0
                .precConv = 0
                .ImporteConv = 0
                Select Case Val(mo_cmbIdFormaPago.BoundText)
                Case 1  'Contado
                Case 2  'SIS
                     If PrSeguro > 0 Then
                        .CantidadSIS = 1
                        .precioSIS = PrSeguro
                        .ImporteSIS = PrSeguro
                        .idestadofacturacion = 10
                        .FechaAutorizaSeguro = Now
                     Else
                        .idTipoFinanciamiento = 1
                        'mo_Atenciones.IdFormaPago = 1
                        'mo_cmbIdFormaPago.BoundText = "1"
                     End If
                Case 3  'Soat
                     If PrSeguro > 0 Then
                        .CantidadSOAT = 1
                        .PrecioSOAT = PrSeguro
                        .ImporteSOAT = PrSeguro
                        .idestadofacturacion = 10
                        .FechaAutorizaSeguro = Now
                     Else
                        .idTipoFinanciamiento = 1
                        'mo_Atenciones.IdFormaPago = 1
                        'mo_cmbIdFormaPago.BoundText = "1"
                     End If
                Case 4  'Convenios
                     If PrSeguro > 0 Then
                        .cantidadConv = 1
                        .precConv = PrSeguro
                        .ImporteConv = PrSeguro
                        .idestadofacturacion = 10
                        .FechaAutorizaSeguro = Now
                     Else
                        .idTipoFinanciamiento = 1
                        'mo_Atenciones.IdFormaPago = 1
                        'mo_cmbIdFormaPago.BoundText = "1"
                     End If
                Case 9  'Exonerados
                    .importeEXO = PrSeguro
                    .idestadofacturacion = 10
                    .idTipoFinanciamiento = 1
                    .FechaAutorizaEXO = Now
                End Select
                '********pone Seguros en forma automatica, sin necesidad de ir a SEGUROS-fin
            End With
            
            Set CargarNuevoFolder = oFolder

End Function


Sub CargaDeServiciosAFacturar()
    
    Set mo_FacturacionServicios = New Collection
    
    If mi_Opcion = sghAgregar Or mi_Opcion = sghModificar Then
        mo_FacturacionServicios.Add CargarTipoDeConsulta()
        
        If Me.chkNuevoCarne.Value = 1 Then
            mo_FacturacionServicios.Add CargarNuevoCarne()
        End If
        
        If Me.chkDuplicadoCarne.Value = 1 Then
            mo_FacturacionServicios.Add CargarDuplicadoCarne()
        End If
        
        If Me.chkNuevoFolder.Value = 1 Then
            mo_FacturacionServicios.Add CargarNuevoFolder()
        End If
        'SCCQ 28-04-2021 Cambio 64 Inicio
        'Validar que sea cuenta SIS y código prestacional 056
        If mo_cmbIdFormaPago.BoundText = 2 And Me.ucSISfuaCodPrestacion1.CodigoPrestacion = "056" Then 'SIS
            'Verificar si se encuentra activo el parámetro 602
            If lcBuscaParametro.SeleccionaFilaParametro(602) = "S" Then
                'Si se encuentra activo, agregar la orden segun valor del servicio en el parámetro 602 ("Oximetría no invasiva para determinar saturación de oxígeno")
                mo_FacturacionServicios.Add CargarProcedimientoFUA(lcBuscaParametro.SeleccionaFilaParametroValorInt(602)) 'Agregar orden para "Oximetría no invasiva para determinar saturación de oxígeno"
            End If
            
            'Verificar si se encuentra activo el parámetro 603
            If lcBuscaParametro.SeleccionaFilaParametro(603) = "S" Then
                'Si se encuentra activo, agregar la orden segun valor del servicio en el parámetro 603("Atención de enfermería")
                mo_FacturacionServicios.Add CargarProcedimientoFUA(lcBuscaParametro.SeleccionaFilaParametroValorInt(603)) 'Agregar orden para "Atención de enfermería"
            End If
        
        End If
        'SCCQ 28-04-2021 Cambio 64 Fin

    End If
        
End Sub

Sub GrabaImagenesEnRutaDelServidor()
    Dim lcArchivoElegido As String
    Dim lcArchivoImagenFinal As String
    lcArchivoElegido = ucPacientesDetalle1.ArchivoElegido
    lcArchivoImagenFinal = wxParametro237 & "\" & Trim(Str(mo_Pacientes.NroHistoriaClinica)) & ".JPG"
    If lcArchivoElegido = "DEL" Then
       Kill lcArchivoImagenFinal
    ElseIf lcArchivoElegido <> "" Then
        pi_imagen.Picture = LoadPicture(lcArchivoElegido)
        SavePicture pi_imagen, lcArchivoImagenFinal
    End If
End Sub


'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------
Function AgregarDatos() As Boolean
    Dim esPacienteNuevo As Boolean
    esPacienteNuevo = False
    If mo_Pacientes.idPaciente = 0 Then
        esPacienteNuevo = True
    End If
        
    AgregarDatos = False
    '5seg/1seg
  
    If mo_AdminAdmision.AdmisionCEAgregar(mo_CuentasAtencion, mo_Atenciones, mo_Pacientes, mo_Cita, mo_Historia, _
                        Me.ucPacientesDetalle1.TipoNumeracionAnterior, mo_Diagnosticos, mo_Procedimientos, _
                        mo_Examenes, Me.ucPacientesDetalle1.IdHistoriaClinicaAnterior, mo_FacturacionServicios, _
                        mo_DOFacturacionPaquetes, lbYaSeTransfirioHCdeUnServicioAotro, mo_lnIdTablaLISTBARITEMS, _
                        mo_lcNombrePc, tabAdmision.Caption, oDoSunasaPacientesHistoricos, _
                        mo_DoAtencionDatosAdicionales, mo_DoPacientesDatosAdd) Then
        txtNroOrdenPago.Text = mo_AdminAdmision.IdOrdenPago
        AgregarDatos = True
    
        If Val(wxParametro208) <> 7686 Then
            GrabaImagenesEnRutaDelServidor
            If esPacienteNuevo = True Then
                Dim o_ReglasIntegracion As New ReglasIntegracion
                Call o_ReglasIntegracion.EnviarDatosPacienteRisPacs(mo_Pacientes)
            End If
        End If
    End If
    ms_MensajeError = mo_AdminAdmision.MensajeError
    If ms_MensajeError = "" Then
        '
        ucPacientesDetalle1.idPaciente = mo_Pacientes.idPaciente 'Actualizado 19092014
        If Val(wxParametro208) <> 7686 Then
            mo_ReglasDeProgMedica.CitasWebActualizaMantCita mo_Cita, sghAgregar    '1seg/1seg
            '
            If mo_Atenciones.IdFormaPago > 1 Then 'no se considera los PAGANTES, porque se espera que vaya a CAJA, allí si se considera este proceso
                Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
                mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar mo_Atenciones.idCuentaAtencion, False, 0   '15seg1seg
                Set mo_ReglasFacturacion = Nothing
            End If
        End If
        If wxParametro302 = "S" And mo_Atenciones.IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS And _
                                                                             lbEncontroAfiliadoEnWebSIS = True Then

           lcTipoFormatoSIS = IIf(Trim(lcTipoFormatoSIS) = "", Trim(Str(sghSIScodigo.sghAfiliacionAUXgratis)), lcTipoFormatoSIS)
           mo_ReglasSISgalenhos.SisFiliacionesActualizarAfiliadoDesdeWEB lcDniSIS, lnAfiliacionSIS1, lnAfiliacionSIS2, _
                                            lnAfiliacionSIS3, lnAfiliacionSIS5, lcTipoFormatoSIS, _
                                            wxParametro323
        End If
    End If
    
    
    
    
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------
Function ModificarDatos() As Boolean
        Dim lbNecesitaCrearDatosEnTablasFacturacion As Boolean
        lbNecesitaCrearDatosEnTablasFacturacion = True
        If mo_Atenciones.IdFormaPago = 1 And txtNroOrdenPago.Text <> "" Then
           'es un Paciente PARTICULAR que YA PAGO o YA TIENE ORDEN DE PAGO
           lbNecesitaCrearDatosEnTablasFacturacion = False
           If lcPagoCita = txtNroOrdenPago.Text Then
               mo_Cita.IdEstadoCita = 4
           End If
        End If
        
        Dim esPacienteNuevo As Boolean
        esPacienteNuevo = False
        If mo_Pacientes.idPaciente = 0 Then
            esPacienteNuevo = True
        End If
        
        ModificarDatos = mo_AdminAdmision.AdmisionCEModificar(mo_CuentasAtencion, mo_Atenciones, mo_Pacientes, mo_Cita, _
                                          mo_Historia, Me.ucPacientesDetalle1.TipoNumeracionAnterior, mo_Diagnosticos, _
                                          mo_Procedimientos, mo_Examenes, Me.ucPacientesDetalle1.IdHistoriaClinicaAnterior, _
                                          mo_FacturacionServicios, mo_FacturacionServiciosPorEliminar, _
                                          mo_DOFacturacionPaquetes, mo_DOFacturacionPaquetesAnt, _
                                          lbYaSeTransfirioHCdeUnServicioAotro, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, _
                                          lc_AntecedentePersonal & tabAdmision.Caption, oDoSunasaPacientesHistoricos, _
                                          mo_DoAtencionDatosAdicionales, lbNecesitaCrearDatosEnTablasFacturacion, _
                                          lbEsUnEPSdesdeAgregarCE, Val(txtNroOrdenPago.Text), mo_DoPacientesDatosAdd)
        ms_MensajeError = mo_AdminAdmision.MensajeError
        If ms_MensajeError = "" Then
            If Val(wxParametro208) <> 7686 Then
                If esPacienteNuevo = True Then
                    Dim o_ReglasIntegracion As New ReglasIntegracion
                    Call o_ReglasIntegracion.EnviarDatosPacienteRisPacs(mo_Pacientes)
                End If
                '
                mo_ReglasDeProgMedica.CitasWebActualizaMantCita mo_Cita, sghModificar
                '
                GrabaImagenesEnRutaDelServidor
                '
                If mo_Atenciones.IdFormaPago > 1 Then     'no se considera los PAGANTES, porque se espera que vaya a CAJA, allí si se considera este proceso
                    Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
                    mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar mo_Atenciones.idCuentaAtencion, False, 0
                    Set mo_ReglasFacturacion = Nothing
                End If
            End If
            If wxParametro302 = "S" And mo_Atenciones.IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
               mo_ReglasSISgalenhos.SisFuaAtencionActualizaDatosDesdeHospEmegCE mo_Atenciones.idCuentaAtencion, _
                                                                      mo_Atenciones.idTipoServicio, mo_Atenciones.idAtencion, _
                                                                      mo_lnIdTablaLISTBARITEMS, ml_idUsuario
            End If
        End If
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------
Function EliminarDatos(oConexion As Connection) As Boolean
    ms_MensajeError = mo_AdminAdmision.VerificaSiTieneMovimientoFarmaciaOservicio(mo_CuentasAtencion.idCuentaAtencion, _
                                                                                  mo_Atenciones.idTipoServicio, oConexion)
    If ms_MensajeError = "" Then
        Dim oEpisodioClinico As EpisodioClinico
        oEpisodioClinico = EpisodioClinicoDevuelveDatos
        '
        mo_CuentasAtencion.idEstado = 9 'anulado
        mo_Atenciones.IdEstadoAtencion = 0  'anulado
        EliminarDatos = mo_AdminAdmision.AdmisionCEAnular(mo_CuentasAtencion, mo_Atenciones, mo_Pacientes, mo_Cita, _
                                                          mo_Historia, Me.ucPacientesDetalle1.TipoNumeracionAnterior, _
                                                          mo_Diagnosticos, mo_Procedimientos, mo_Examenes, _
                                                          Me.ucPacientesDetalle1.IdHistoriaClinicaAnterior, _
                                                          mo_FacturacionServicios, mo_FacturacionServiciosPorEliminar, _
                                                          mo_DOFacturacionPaquetesAnt, mo_lnIdTablaLISTBARITEMS, _
                                                          mo_lcNombrePc, tabAdmision.Caption, oEpisodioClinico)
        ms_MensajeError = mo_AdminAdmision.MensajeError
        If ms_MensajeError <> "" Then
           MsgBox ms_MensajeError
        Else
            '
            mo_ReglasDeProgMedica.CitasWebActualizaMantCita mo_Cita, sghEliminar
            '
            If GrabaAtencionJamo = True Then
               GrabaAtencionPerinatal
            End If
    '        EliminarDatos = mo_AdminAdmision.RecetaEliminar(mo_Atenciones.idCuentaAtencion, mo_Atenciones.IdServicioIngreso, ml_idUsuario, _
    '                                         lnRecetaRayosX, lnRecetaEcografiaO, lnRecetaEcografiaG, lnRecetaTomografia, _
    '                                         lnRecetaAnatomiaP, lnRecetaPatologiaC, lnRecetaBancoS, lnRecetaFarmacia, _
    '                                         Me.ucRecetas1.DevuelveRayosX, Me.ucRecetas1.DevuelveEcografiaO, _
    '                                         Me.ucRecetas1.DevuelveEcografiaG, Me.ucRecetas1.DevuelveTomografia, _
    '                                         Me.ucRecetas1.DevuelveAnatomia, Me.ucRecetas1.DevuelvePatologia, _
    '                                         Me.ucRecetas1.DevuelveBancoSangre, Me.ucRecetas1.DevuelveFarmacia, _
    '                                         mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "Paciente : " & ms_NombrePaciente)
        End If
    Else
        MsgBox ms_MensajeError & Chr(13) & "La Anulación tendrá que realizarlo FACTURACION ", vbInformation, "Consulta externa"
    End If
End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla CuentasAtencion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlosControles()
        Dim oRecetaCabecera As New RecetaCabecera
        Dim oRsCabeceraReceta As New Recordset
        Dim oConexion As New Connection
        Dim lcFormaCita As String
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighEntidades.CadenaConexion
        '
        btnAceptar.Enabled = True
        '
        If oRsFormaPago.State = adStateOpen Then oRsFormaPago.Close
        Set oRsFormaPago = mo_AdminServiciosComunes.TiposFinanciamientoSegunFiltro("esFuenteFinanciamiento=1")
        Set mo_cmbIdFormaPago.RowSource = oRsFormaPago
        mo_cmbIdFormaPago.ListField = "Descripcion"
        mo_cmbIdFormaPago.BoundColumn = "idTipoFinanciamiento"
        '
        '1ro:   CARGAR DATOS DE LA CITA
        CargarDatosDeCitasALosControles oConexion
       
        '2do:   CARGAR DATOS DE LA ATENCION
        CargarDatosDelaAtencion oConexion
                
        
        '4to:   PARA VISUALIZAR LA UBICACION DEL PACIENTE AL DIA DE LA ATENCION
        mo_DoUbicacionPaciente.DireccionDomicilio = mo_DoAtencionDatosAdicionales.DireccionDomicilio
        Me.ucPacientesDetalle1.ReemplazarDatosDeUbicacion mo_DoUbicacionPaciente
        
        '5to:   CARGAR DATOS DE LOS DIAGNOSTICOS POR ATENCION
               
        'Cargar datos de servicios
        CargarDatosDeServiciosFacturados Me.idAtencion, mo_Cita.idProducto, oConexion
        
        'Carga datos de Triaje y Atencion CE (debb-jamo)
       
        'Verifica la Cta Atencion
        'Dim mo_CuentasAtencion As New DOCuentaAtencion
        If lbCargaAlaVezCitaPacienteAtencionDA = False Then
           Set mo_CuentasAtencion = mo_AdminFacturacion.CuentasAtencionSeleccionarPorId(Me.idCuentaAtencion, oConexion)
        End If
        lblEstadoCta = mo_ReglasFarmacia.DevuelveEstadoActualDeEstadoCuenta("idEstado=" & mo_CuentasAtencion.idEstado, oConexion)
        If mo_CuentasAtencion.idEstado <> 1 Then btnAceptar.Enabled = False
        txtNroCuenta.Text = mo_CuentasAtencion.idCuentaAtencion
        ms_MensajeError = VerSiTieneServicioAutomaticoPorEstancia(oConexion)
        'Ya tuvo movimientos(Farmacia/servicios), no podrá cambiar de plan
        If mi_Opcion = sghModificar Then
            ms_MensajeError = mo_AdminAdmision.VerificaSiTieneMovimientoFarmaciaOservicio(mo_Atenciones.idCuentaAtencion, mo_Atenciones.idTipoServicio, oConexion)
            If ms_MensajeError <> "" Then
               mo_Formulario.HabilitarDeshabilitar Me.cmbFuenteFinanciamiento, False
               Me.ucMensajeParpadeando1.MensajeDeTexto = ms_MensajeError
               Me.ucMensajeParpadeando1.Visible = True
            End If
            ms_MensajeError = ""
        Else
           ucMensajeParpadeando1.MensajeDeTexto = ""
           ucMensajeParpadeando1.Visible = False
        End If
        
        '
        DeudasPendientesDeAnterioresAtenciones oConexion
        '
        PaqueteCargaDatos mo_Atenciones.idCuentaAtencion, oConexion
        '
        If mo_Atenciones.idSunasaPacienteHistorico > 0 Then
            If mo_AdminFacturacion.TiposFinanciamientoGeneraReciboPago(Val(mo_cmbIdFormaPago.BoundText), oConexion) = True Then
                Me.UcPacientesSunasa1.YaNoTieneSeguro
                Me.UcPacientesSunasa1.HabilitaFrame False
            Else
                Me.UcPacientesSunasa1.HabilitaFrame True
                Me.UcPacientesSunasa1.idSunasaPacienteHistorico = mo_Atenciones.idSunasaPacienteHistorico
                Me.UcPacientesSunasa1.CargarDatosPorId
            End If
        End If
        UcPacientesSunasa1.idTipoFinanciamiento = Val(mo_cmbIdFormaPago.BoundText)
        '
        mb_NecesitaTriaje = mo_AdminAdmision.ElServicioNecesitaTriaje(mo_Atenciones.IdServicioIngreso, oConexion, lbElConsultorioUsaModuloPerinatal, lbElConsultorioUsaModuloMaterno)
        '
        '
        lnAfiliacionSIS4 = mo_DoAtencionDatosAdicionales.idSiaSis
        lcSIScodigo = mo_DoAtencionDatosAdicionales.SisCodigo
        '
        Me.UcEpisodioClinico1.idPaciente = mo_Atenciones.idPaciente
        Me.UcEpisodioClinico1.idAtencion = mo_Atenciones.idAtencion
        Me.UcEpisodioClinico1.Inicializar
        Me.UcEpisodioClinico1.Limpiar
        Me.UcEpisodioClinico1.CargaEpisodiosHistoricos
        Me.UcEpisodioClinico1.CargarDatosAlosControles oConexion
        '
        Me.FraGeneraCita.Enabled = True
        lcFormaCita = mo_AdminAdmision.CitasFormaCitaSeleccionaXcuenta(mo_Atenciones.idCuentaAtencion, oConexion)
        Select Case lcFormaCita
        Case "N"
             Me.optCNorma.Value = True
        Case "W"
             Me.optCweb.Value = True
             Me.FraGeneraCita.Enabled = False
        Case "T"
             Me.optCtelefono.Value = True
        End Select
        '
        oConexion.Close
        Set oConexion = Nothing
        Set oRecetaCabecera = Nothing
        Set oRsCabeceraReceta = Nothing
        '
        If mo_Cita.IdEstadoCita = 4 And mo_Atenciones.IdFormaPago = 1 Then
            txtNroOrdenPago.Text = lcPagoCita
            If Not (mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE And mi_Opcion = sghModificar) Then
                MsgBox "La cita está PAGADA, no podrá Modificar/Eliminar hasta que ANULE LA BOLETA", vbInformation, Me.Caption
                Me.btnAceptar.Enabled = False
            End If
        End If
        '
        If mi_Opcion = sghModificar Then
            mo_Formulario.HabilitarDeshabilitar cmbIdServicio, False
            If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE Then
               mo_Formulario.HabilitarDeshabilitar cmbFuenteFinanciamiento, False
               mo_Formulario.HabilitarDeshabilitar Me.cmbFormaPago, False
               mo_Formulario.HabilitarDeshabilitar fraDatosCita, False
               mo_Formulario.HabilitarDeshabilitar fraDatosReferenciaOrigen, False
            End If
        End If
        '
        HaceVisibleOnoBotonFUA
        '
        
        '
        btnImprimeFiliacion.Enabled = True
End Sub




Sub PaqueteCargaDatos(lnIdCuentaAtencion As Long, oConexion As Connection)
    Dim oRsTmp As New Recordset
    Dim lcSql As String
    Set oRsTmp = mo_AdminFacturacion.FacturacionPaquetesSeleccionarPorid("dbo.FacturacionPaquetes.AtencionId=" & Trim(Str(lnIdCuentaAtencion)))
    
    If oRsTmp.RecordCount > 0 Then
       mo_DOFacturacionPaquetes.IdComprobantePago = oRsTmp.Fields!IdComprobantePago
       mo_DOFacturacionPaquetes.IdOrdenPago = oRsTmp.Fields!IdOrdenPago
       mo_DOFacturacionPaquetes.idProducto = oRsTmp.Fields!idProducto
       mo_DOFacturacionPaquetesAnt.IdComprobantePago = oRsTmp.Fields!IdComprobantePago
       mo_DOFacturacionPaquetesAnt.IdOrdenPago = oRsTmp.Fields!IdOrdenPago
       mo_DOFacturacionPaquetesAnt.idProducto = oRsTmp.Fields!idProducto
       Me.txtNserie.Text = oRsTmp.Fields!nroSerie
       Me.txtNboleta.Text = oRsTmp.Fields!nrodocumento
       lcNserieAnt = oRsTmp.Fields!nroSerie
       lcNboletaAnt = oRsTmp.Fields!nrodocumento
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
End Sub

Sub DeudasPendientesDeAnterioresAtenciones(oConexion As Connection)
        'Deudas
        If mo_lnIdTablaLISTBARITEMS <> sghOpcionGalenHos.sghRegistroAtencionCE Then   'Registro de Atención
            ms_MensajeError = mo_AdminFacturacion.DevuelveDeudaPacienteDeAntencionesAnteriores(ml_IdPaciente, oConexion, ml_idCuentaAtencion)
            If ms_MensajeError <> "" Then
               ucMensajeParpadeando1.MensajeDeTexto = "Deudas:   " & ms_MensajeError
               ucMensajeParpadeando1.Visible = True
               'debb-29/02/2016 (inicio)
               If mi_Opcion = sghAgregar And InStr(ms_MensajeError, "<FALLECIDO>") > 0 Then
                  btnCancelar_Click
               End If
               'debb-29/02/2016 (fin)
            Else
                ucMensajeParpadeando1.MensajeDeTexto = ""
                ucMensajeParpadeando1.Visible = False
            End If
        End If
        ms_MensajeError = ""
End Sub

Sub CargarDatosDeServiciosFacturados(lIdAtencion As Long, lIdTipoConsulta As Long, oConexion As Connection)
Dim oRecordset As New Recordset
Dim oTabla As New Parametros
Dim lcSql As String
    Set oTabla.Conexion = oConexion
    '
    Set oRecordset = mo_AdminFacturacion.ServiciosFacturadosPorIdAtencion(lIdAtencion)
    '
    oRecordset.Filter = "idProducto=" & lIdTipoConsulta
    If Not (oRecordset.EOF And oRecordset.BOF) Then
        lnIdFactServicios = oRecordset.Fields!IdOrden
        Me.cmbIdTipoConsulta.Tag = lIdTipoConsulta
        mo_Formulario.HabilitarDeshabilitar Me.cmbIdTipoConsulta, (oRecordset!idestadofacturacion = sghEstadoFacturacion.sghRegistrado)
    End If
    oRecordset.Filter = "idProducto=" & oTabla.ObtenerCodigoDeDuplicadoCarnet()
    If Not (oRecordset.EOF And oRecordset.BOF) Then
        Me.chkDuplicadoCarne.Tag = 1
        Me.chkDuplicadoCarne.Value = 1
        mo_Formulario.HabilitarDeshabilitar Me.chkDuplicadoCarne, (oRecordset!idestadofacturacion = sghEstadoFacturacion.sghRegistrado)
    End If
    oRecordset.Filter = "idProducto=" & oTabla.ObtenerCodigoDeNuevoCarnet()
    If Not (oRecordset.EOF And oRecordset.BOF) Then
        Me.chkNuevoCarne.Tag = 1
        Me.chkNuevoCarne.Value = 1
        mo_Formulario.HabilitarDeshabilitar Me.chkNuevoCarne, oRecordset!idestadofacturacion = sghEstadoFacturacion.sghRegistrado
        Me.chkNuevoCarne.Visible = True
    Else
        Me.chkNuevoCarne.Visible = False
    End If
    oRecordset.Filter = "idProducto=" & oTabla.ObtenerCodigoDeFolder()
    If Not (oRecordset.EOF And oRecordset.BOF) Then
        Me.chkNuevoFolder.Value = 1
        Me.chkNuevoFolder.Tag = 1
        mo_Formulario.HabilitarDeshabilitar Me.chkNuevoFolder, oRecordset!idestadofacturacion = sghEstadoFacturacion.sghRegistrado
        Me.chkNuevoFolder.Visible = True
    Else
        Me.chkNuevoFolder.Visible = False
    End If
    oRecordset.Close
    Set oRecordset = Nothing
    Set oTabla = Nothing
End Sub

Sub CargarDatosDelaAtencion(oConexion As Connection)
Dim oDoMedico As New DOMedico
Dim oDOEmpleado As New dOEmpleado
Dim oDOEspecialidades As New Collection
Dim lcEstadoAtencion As String
        'El Id de atencion se obtuvo al momento de cargar los datos de la cita
        If lbCargaAlaVezCitaPacienteAtencionDA = False Then
           Set mo_Atenciones = mo_AdminAdmision.AtencionesSeleccionarPorId(Me.idAtencion, oConexion)
        End If
        If mo_AdminAdmision.MensajeError <> "" Then
             MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminServiciosComunes.MensajeError, vbInformation, Me.Caption"
             mb_ExistenDatos = False
             Exit Sub
        End If
        lblNroAtencion.Caption = "N° Atención: " & Trim(Str(Me.idAtencion))
        If Not mo_Atenciones Is Nothing Then
           With mo_Atenciones
                
                Me.idMedico = .IdMedicoIngreso
                Me.idCuentaAtencion = .idCuentaAtencion
                'Carga de la especialidad
                mo_cmbIdEspecialidadMedico.BoundColumn = "IdEspecialidad"
                mo_cmbIdEspecialidadMedico.ListField = "DescripcionLarga"
                Dim rsEspecialidad As New Recordset
                Set rsEspecialidad = mo_AdminServiciosHosp.EspecialidadesSeleccionarporMedico(.IdMedicoIngreso, oConexion)
                Set mo_cmbIdEspecialidadMedico.RowSource = rsEspecialidad
                If rsEspecialidad.RecordCount = 1 Then
                    mo_Formulario.HabilitarDeshabilitar cmbIdEspecialidadMedico, False
                End If
                
                'Carga datos del medico
                If mo_ReglasDeProgMedica.MedicosSeleccionarPorId(.IdMedicoIngreso, oDoMedico, oDOEmpleado, oDOEspecialidades, oConexion) Then
                    Me.txtMedico.Text = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres + " (" & oDOEmpleado.CodigoPlanilla & ")"
                Else
                    Me.txtMedico.Text = ""
                End If
                mo_cmbIdEspecialidadMedico.BoundText = .IdEspecialidadMedico
                cmbIdEspecialidadMedico_LostFocus
                mo_cmbIdTipoConsulta.BoundText = mo_Cita.idProducto
    
                'Carga datos del consultorios
                mo_cmbIdServicio.BoundColumn = "IdServicio"
                mo_cmbIdServicio.ListField = "DescripcionLarga"
                
                Set rsServicio = mo_AdminServiciosHosp.ServiciosSeleccionarConsultoriosPorEspecialidaddebb(Val(mo_cmbIdEspecialidadMedico.BoundText), sghFiltraAnuladosYactivos, oConexion)
                Set mo_cmbIdServicio.RowSource = rsServicio
                
                
                'Carga datos de la atención
                mo_cmbIdViasAdmision.BoundText = .IdOrigenAtencion
                Me.txtHoraInicio.Text = IIf(.HoraIngreso = "", sighEntidades.HORA_VACIA_HM, .HoraIngreso)
                Me.txtFechaIngreso.Text = IIf(.FechaIngreso = 0, sighEntidades.FECHA_VACIA_DMY, .FechaIngreso)
                mo_cmbIdTipoServicio.BoundText = .idTipoServicio
                Me.txtEdadEnDias.Text = .Edad
                Me.txtEdadEnDias.Tag = .Edad
                
                mo_cmbIdTipoEdad.BoundText = .idTipoEdad
                cmbIdTipoEdad.Tag = .idTipoEdad
                
                mo_cmbIdServicio.BoundText = .IdServicioIngreso
                mo_cmbIdFormaPago.BoundText = .IdFormaPago
                mo_cmbIdFuentesFinanciamiento.BoundText = .IdFuenteFinanciamiento
                lnFormaPagoAnterior = .IdFormaPago
                Select Case .IdEstadoAtencion
                Case 0
                    lcEstadoAtencion = "Anulado"
                    btnAceptar.Enabled = False
                Case 1
                    lcEstadoAtencion = "Registrado"
                Case 2
                    lcEstadoAtencion = "Cerrado"
                    btnAceptar.Enabled = False
                End Select
                '
                lbEsUnEPSdesdeAgregarCE = False
                If .EpsPorcentaje > 0 Then
                   ucEPS1.Porcentaje = .EpsPorcentaje
                   ucEPS1.Visible = True
                   lbEsUnEPSdesdeAgregarCE = True
                End If
                '
                mb_ExistenDatos = True
           End With
           '
           If lbCargaAlaVezCitaPacienteAtencionDA = False Then
              Set mo_DoAtencionDatosAdicionales = mo_AdminAdmision.AtencionesDatosAdicionalesSeleccionarPorId(Me.idAtencion, oConexion)
           End If
           With mo_DoAtencionDatosAdicionales
                mo_cmbIdTipoReferenciaOrigen.BoundText = .IdTipoReferenciaOrigen
                CompletarDatosDelEstablecimientoEnElLoad .IdEstablecimientoOrigen, .IdEstablecimientoNoMinsaOrigen, txtIdEstablecimientoOrigen, txtNombreOrigenReferencia, .IdTipoReferenciaOrigen
                txtReferenciaO.Text = .NroReferenciaOrigen
                'debb-21/06/2016 (inicio)
                Dim lcDxCodigo As String, lcDx As String
                PVcomboBoxUbicaPosicion .referenciaOservicio, cmbServicioReferenciaO
                txtDxReferencia.Tag = .referenciaOidDiagnostico
                mo_AdminServiciosComunes.DiagnosticosSeleccionarPorIdDevuelveDescripcion Val(txtDxReferencia.Tag), _
                                                                                         oConexion, lcDxCodigo, lcDx
                txtDxReferencia.Text = lcDxCodigo
                lblDxReferencia1.Text = lcDx
                'debb-21/06/2016 (fin)
                'franklin 2017
                Me.txtMedicoRef.Text = .ReferenciaMedicoOColeg
                BuscaMedicoRerencia .ReferenciaMedicoOIdcolegio
                
           End With
           'ESTOS DATOS SE UTILIZARAN MAS ADELANTE PARA ACTUALIZAR LA UBICACION DE PACIENTE
           Dim oPacientesTmp As New SIGHComun.doPaciente
           Set oPacientesTmp = mo_paciente
           If Not oPacientesTmp Is Nothing Then
                With oPacientesTmp
                     lcHistoriaYpaciente = "(" & Trim(Str(oPacientesTmp.NroHistoriaClinica)) & ") " & Trim(oPacientesTmp.ApellidoPaterno) & " " & Trim(oPacientesTmp.ApellidoMaterno) & " " & Trim(oPacientesTmp.PrimerNombre)
                     Me.Caption = Me.Caption & " (HC: " & Trim(oPacientesTmp.NroHistoriaClinica) & " " & _
                                  Trim(oPacientesTmp.ApellidoPaterno) & " " & Trim(oPacientesTmp.ApellidoMaterno) & _
                                  " " & Trim(oPacientesTmp.PrimerNombre) & ") (Estado: " & lcEstadoAtencion & _
                                  ")(Edad: " & Trim(txtEdadEnDias.Text) & " " & Left(Trim(cmbIdTipoEdad.Text), 1) & _
                                  ")(Gs: " & IIf(IsNull(oPacientesTmp.GrupoSanguineo), "", oPacientesTmp.GrupoSanguineo) & _
                                  ", Frh: " & IIf(IsNull(oPacientesTmp.FactorRh), "", oPacientesTmp.FactorRh) & ")"
                     mo_DoUbicacionPaciente.IdPaisDomicilio = .IdPaisDomicilio
                     mo_DoUbicacionPaciente.IdCentroPobladoDomicilio = .IdCentroPobladoDomicilio
                     
                     mo_DoUbicacionPaciente.IdPaisProcedencia = .IdPaisProcedencia
                     mo_DoUbicacionPaciente.IdCentroPobladoProcedencia = .IdCentroPobladoProcedencia
                     
                     mo_DoUbicacionPaciente.DireccionDomicilio = .DireccionDomicilio
                End With
                Me.ucPacientesDetalle1.CargarDatosDePacienteALosControlesSinBuscar oPacientesTmp, wxParametro242, wxParametro287
           End If
           '
           Set oPacientesTmp = Nothing
       Else
           mb_ExistenDatos = False
           Exit Sub
       End If
       Set oDoMedico = Nothing
       Set oDOEmpleado = Nothing
       Set oDOEspecialidades = Nothing
End Sub



Sub CargarDatosDeCitasALosControles(oConexion As Connection)
    
    Set mo_Cita = New DOCita
    Me.idAtencion = 0
    If lbCargaAlaVezCitaPacienteAtencionDA = False Then
       mb_ExistenDatos = mo_AdminAdmision.CitasSeleccionarPorId(ml_IdCita, mo_Cita, mo_paciente, oConexion)
    Else
       mo_Cita.IdCita = ml_IdCita
       mb_ExistenDatos = mo_AdminAdmision.AtencionesPacientesCitasDatosadicionalesSeleccionarPorId(mo_paciente, _
                                           mo_Atenciones, mo_DoAtencionDatosAdicionales, _
                                           oConexion, mo_CuentasAtencion, True, mo_Cita)
    End If
    If mo_AdminAdmision.MensajeError <> "" Then
         MsgBox "No se pudo obtener los datos" + Chr(13) + mo_AdminAdmision.MensajeError + Chr(13) + Chr(13) + "Salga del Sistema a Windows y vuelva a ingresar", vbInformation, Me.Caption
         mb_ExistenDatos = False
         Me.Visible = False
         Exit Sub
    End If
       
    If mb_ExistenDatos Then
         With mo_Cita
             Me.IdCita = .IdCita
             ml_idAtencion = .idAtencion  'IMPORTANTE!!! Carga el IdAtencion
             Me.idPaciente = .idPaciente
             Me.IdProgramacion = .IdProgramacion
             Me.IdEstadoCita = .IdEstadoCita
             Me.txtFechaIngreso.Text = .fecha
             Me.txtHoraInicio.Text = .HoraInicio
             Me.txtHoraFin.Text = .HoraFin
             mo_lbEsCitaAdicional = .EsCitaAdicional
             mb_ExistenDatos = True
         End With
         
    Else
        Me.idAtencion = 0
        mb_ExistenDatos = False
        Me.Visible = False
        Exit Sub
    End If
   
End Sub


Sub LimpiarFormulario()

           'LIMPIAR DATOS DE LA CUENTA DE ATENCION
           Me.idCuentaAtencion = 0
           Me.idAtencion = 0
           
           'LIMPIAR DATOS DE LA ATENCION
           mo_cmbIdTipoReferenciaOrigen.BoundText = ""
           Me.txtIdEstablecimientoOrigen.Text = ""
           Me.txtEdadEnDias.Text = ""
           
           Me.ucPacientesDetalle1.LimpiarDatosDePaciente wxParametro211, ldFechaActualServidor
            
                      
End Sub

Private Sub cmbIdTipoReferenciaOrigen_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoReferenciaOrigen
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdTipoReferenciaOrigen_LostFocus()
   If cmbIdTipoReferenciaOrigen.Text <> "" Then
       mo_cmbIdTipoReferenciaOrigen.BoundText = Val(Split(cmbIdTipoReferenciaOrigen.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdTipoReferenciaOrigen
End Sub

Private Sub cmbIdTipoReferenciaOrigen_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtIdEstablecimientoOrigen_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdEstablecimientoOrigen
    If KeyCode = vbKeyF1 Then
        btnBuscarEstablecimiento_Click
    End If
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdEstablecimientoOrigen_LostFocus()
   
    CompletarDatosDelEstablecimientoEnElLostFocus txtIdEstablecimientoOrigen, txtNombreOrigenReferencia, Val(mo_cmbIdTipoReferenciaOrigen.BoundText)
    mo_Formulario.MarcarComoVacio txtIdEstablecimientoOrigen
End Sub

Private Sub txtIdEstablecimientoOrigen_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdTipoServicio_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoServicio
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdTipoServicio_LostFocus()
   If cmbIdTipoServicio.Text <> "" Then
       mo_cmbIdTipoServicio.BoundText = Val(Split(cmbIdTipoServicio.Text, " = ")(0))
   End If
   If Not cmbIdTipoServicio.Locked Then mo_Formulario.MarcarComoVacio cmbIdTipoServicio
End Sub

Private Sub cmbIdTipoServicio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtEdadEnDias_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtEdadEnDias
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtEdadEnDias_LostFocus()
   mo_Formulario.MarcarComoVacio txtEdadEnDias
End Sub

Private Sub txtEdadEnDias_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtNserie_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNserie
    AdministrarKeyPreview KeyCode

End Sub














Private Sub txtPrimerNombreBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtSegundoNombreBusqueda.SetFocus
   'mo_Teclado.RealizarNavegacion KeyCode, txtPrimerNombreBusqueda
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtPrimerNombreBusqueda_LostFocus()
   txtPrimerNombreBusqueda.Text = mo_Teclado.CapitalizarNombres(txtPrimerNombreBusqueda.Text)
   If Len(txtPrimerNombreBusqueda.Text) > 0 Then
      btnBuscarPaciente_Click
   End If
End Sub

Private Sub txtPrimerNombreBusqueda_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub




Private Sub txtReferenciaO_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtReferenciaO
End Sub

Private Sub txtSegundoNombreBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
   If UcSISafiliacion1.Visible = False Then
        If Me.grdPacientesEncontrados.Visible = True Then
            grdPacientesEncontrados.SetFocus
        Else
            mo_Teclado.RealizarNavegacion KeyCode, txtSegundoNombreBusqueda
        End If
   Else
        mo_Teclado.RealizarNavegacion KeyCode, txtSegundoNombreBusqueda
   End If
   AdministrarKeyPreview KeyCode
End Sub


Private Sub txtSegundoNombreBusqueda_LostFocus()
   txtSegundoNombreBusqueda.Text = mo_Teclado.CapitalizarNombres(txtSegundoNombreBusqueda.Text)
End Sub

Private Sub txtSegundoNombreBusqueda_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

























Private Sub ucPacientesDetalle1_SeModificoFechaNacimiento(sFechaNacimiento As String, sHoraNacimiento As String)
    
    On Error Resume Next
    Me.txtEdadEnDias = ""
    
    Dim oEdad As Edad
    oEdad = CalcularEdad(CDate(sFechaNacimiento & " " & sHoraNacimiento), CDate(txtFechaIngreso.Text & " " & txtHoraInicio.Text))
    Me.txtEdadEnDias = oEdad.Edad
    mo_cmbIdTipoEdad.BoundText = oEdad.TipoEdad
    
    If Me.txtEdadEnDias.Text = "" Then
        Me.txtEdadEnDias.Text = Me.txtEdadEnDias.Tag
    End If


End Sub

Private Sub ucPacientesDetalle1_SeModificoPacienteNoIdentificado(bPacienteNoIdentificado As Boolean)
    
    If bPacienteNoIdentificado Then
        chkPacienteNuevo.Value = 1
        chkPacienteNuevo.Enabled = False
    Else
        chkPacienteNuevo.Enabled = True
        chkPacienteNuevo.Value = 1
    End If

End Sub



Private Sub ucPacientesDetalle1_SePresionoTeclaEspecial(KeyCode As Integer)
    AdministrarKeyPreview KeyCode
End Sub


Sub CompletarDatosDeEstablecimiento(txtIdEstablecimiento As TextBox, lblNombreEstablecimiento As TextBox, lTipoReferencia As Long)
    
    If lTipoReferencia = 1 Then
        'Dim oBusqueda As New EstablecimientosBusqueda
        Dim oBusqueda As New SIGHNegocios.BuscaEstablecimientos
        Dim oDoEstablecimiento As New DOEstablecimiento
        'oBusqueda.Show 1
        oBusqueda.MostrarFormulario
        If oBusqueda.BotonPresionado = sghAceptar Then
        
            Set oDoEstablecimiento = mo_AdminServiciosComunes.EstablecimientosSeleccionarPorId(oBusqueda.idRegistroSeleccionado)
            If Not oDoEstablecimiento Is Nothing Then
                txtIdEstablecimiento.Tag = oDoEstablecimiento.IdEstablecimiento
                txtIdEstablecimiento.Text = oDoEstablecimiento.Codigo
                lblNombreEstablecimiento = oDoEstablecimiento.nombre
            Else
                txtIdEstablecimiento.Tag = ""
                txtIdEstablecimiento.Text = ""
                lblNombreEstablecimiento = ""
            End If
        End If
        Set oBusqueda = Nothing
        Set oDoEstablecimiento = Nothing
    Else
        'Dim oBusquedaNM As New EstablecimientosNoMinsaBusqueda
        Dim oBusquedaNM As New SIGHNegocios.BuscaEstablecNoMinsa
        Dim oDoEstablecimientoNM As New DOEstablecimientoNoMinsa
        oBusquedaNM.lcNombrePc = mo_lcNombrePc
        oBusquedaNM.idUsuario = ml_idUsuario
        'oBusquedaNM.Show 1
        oBusquedaNM.MostrarFormulario
        If oBusquedaNM.BotonPresionado = sghAceptar Then
            Set oDoEstablecimientoNM = mo_AdminServiciosComunes.EstablecimientosNoMinsaSeleccionarPorId(oBusquedaNM.idRegistroSeleccionado)
            If Not oDoEstablecimientoNM Is Nothing Then
                txtIdEstablecimiento.Tag = oDoEstablecimientoNM.IdEstablecimientoNoMINSA
                txtIdEstablecimiento.Text = oDoEstablecimientoNM.IdEstablecimientoNoMINSA
                lblNombreEstablecimiento = oDoEstablecimientoNM.nombre
            Else
                txtIdEstablecimiento.Tag = ""
                txtIdEstablecimiento.Text = ""
                lblNombreEstablecimiento = ""
            End If
        End If
        Set oBusquedaNM = Nothing
        Set oDoEstablecimientoNM = Nothing
    End If

End Sub
Sub CompletarDatosDelEstablecimientoEnElLostFocus(txtIdEstablecimiento As TextBox, lblNombreEstablecimiento As TextBox, lTipoReferencia As Long)
    
    If txtIdEstablecimiento <> "" Then
        If lTipoReferencia = 1 Then
                Dim oDoEstablecimiento As New DOEstablecimiento
                If mo_AdminServiciosComunes.EstablecimientosSeleccionarPorCodigo(txtIdEstablecimiento.Text, oDoEstablecimiento) Then
                    txtIdEstablecimiento.Tag = oDoEstablecimiento.IdEstablecimiento
                    txtIdEstablecimiento.Text = oDoEstablecimiento.Codigo
                    lblNombreEstablecimiento = oDoEstablecimiento.nombre
                Else
                    txtIdEstablecimiento.Tag = ""
                    txtIdEstablecimiento = ""
                    lblNombreEstablecimiento = ""
                End If
        Else
                Dim oDOEstablecimientoNoMinsa As New DOEstablecimientoNoMinsa
                Set oDOEstablecimientoNoMinsa = mo_AdminServiciosComunes.EstablecimientosNoMinsaSeleccionarPorId(txtIdEstablecimiento.Text)
                If Not oDOEstablecimientoNoMinsa Is Nothing Then
                    txtIdEstablecimiento.Tag = oDOEstablecimientoNoMinsa.IdEstablecimientoNoMINSA
                    txtIdEstablecimiento.Text = oDOEstablecimientoNoMinsa.IdEstablecimientoNoMINSA
                    lblNombreEstablecimiento = oDOEstablecimientoNoMinsa.nombre
                Else
                    txtIdEstablecimiento.Tag = ""
                    txtIdEstablecimiento = ""
                    lblNombreEstablecimiento = ""
                End If
        End If
    End If

End Sub
Sub CompletarDatosDelEstablecimientoEnElLoad(lIdEstablecimiento As Long, lIdEstablecimientoNoMinsa As Long, txtIdEstablecimiento As TextBox, lblNombreEstablecimiento As TextBox, lTipoReferencia As Long)
                
    If lTipoReferencia = 1 Then
        Dim oDoEstablecimiento As New DOEstablecimiento
         Set oDoEstablecimiento = mo_AdminServiciosComunes.EstablecimientosSeleccionarPorId(lIdEstablecimiento)
         If Not oDoEstablecimiento Is Nothing Then
             txtIdEstablecimiento.Text = oDoEstablecimiento.Codigo
             txtIdEstablecimiento.Tag = oDoEstablecimiento.IdEstablecimiento
             lblNombreEstablecimiento = oDoEstablecimiento.nombre
        Else
             txtIdEstablecimiento.Text = ""
             txtIdEstablecimiento.Tag = ""
             lblNombreEstablecimiento = ""
        End If
    Else
        Dim oDOEstablecimientoNoMinsa As New DOEstablecimientoNoMinsa
         Set oDOEstablecimientoNoMinsa = mo_AdminServiciosComunes.EstablecimientosNoMinsaSeleccionarPorId(lIdEstablecimientoNoMinsa)
         If Not oDOEstablecimientoNoMinsa Is Nothing Then
             txtIdEstablecimiento.Text = oDOEstablecimientoNoMinsa.IdEstablecimientoNoMINSA
             txtIdEstablecimiento.Tag = oDOEstablecimientoNoMinsa.IdEstablecimientoNoMINSA
             lblNombreEstablecimiento = oDOEstablecimientoNoMinsa.nombre
        Else
             txtIdEstablecimiento.Text = ""
             txtIdEstablecimiento.Tag = ""
             lblNombreEstablecimiento = ""
         End If
    End If

End Sub


Sub MuestraCitasAnteriores(oConexion As Connection, lbSoloMuestraGrillaCitasAnteriores As Boolean)
   Dim lcSql As String
   Dim oRsTmp As New ADODB.Recordset
   Set oRsTmp = mo_AdminFacturacion.CitasMuestraAteriores(idPaciente, oConexion)
   oRsTmp.Filter = "fecha> " & ldFechaActualServidor
   lcSql = Trim(Str(oRsTmp.RecordCount))
   oRsTmp.Filter = ""
   Set oRsTmp.ActiveConnection = Nothing
   Set grdAnteriores.DataSource = oRsTmp
   If lbSoloMuestraGrillaCitasAnteriores = False Then
        If Val(lcSql) > 0 Then
           If mi_nroHistoriaCitadoXmedico = 0 Then
           If MsgBox("      Existen " & lcSql & " CITAS mayores a HOY del PACIENTE elegido        " & Chr(13) & _
                     "      (color VERDE en la LISTA superior derecha de la ventana actual)        " & Chr(13) & Chr(13) & _
                     "                        Esta seguro proseguir ?                             ", vbQuestion + vbYesNo, "Mensaje") = vbNo Then
              btnLimpiar_Click
           End If
           End If
        End If
        
        '
        UcPacientesSunasa1.idPaciente = ml_IdPaciente
        UcPacientesSunasa1.CargarDatosDelUltimoSeguroDelPacienteALosControles oConexion
        '
        DeudasPendientesDeAnterioresAtenciones oConexion
   End If
End Sub


Sub ImprimePreCuenta()
    Dim oReporte As New RptCaja
    Dim lcPaciente As String
    Dim lcMedico As String
    Dim lcCola As String
    Dim oRsTmp3 As New Recordset
    If mi_Opcion <> sghAgregar Then
       Me.ucPacientesDetalle1.CargarDatosAlObjetoDatos mo_Pacientes, mo_Historia, mo_DoPacientesDatosAdd
    End If
    lcPaciente = Trim(mo_Pacientes.ApellidoPaterno) & " " & Trim(mo_Pacientes.ApellidoMaterno) & " " & Trim(mo_Pacientes.PrimerNombre)
    If mo_Pacientes.SegundoNombre <> "" Then
       lcPaciente = lcPaciente & " " & Trim(mo_Pacientes.SegundoNombre)
    End If
    If mo_Pacientes.TercerNombre <> "" Then
      lcPaciente = lcPaciente & " " & Trim(mo_Pacientes.TercerNombre)
    End If
    lcMedico = Left(txtMedico.Text, InStr(txtMedico.Text, "(") - 1)
    If Val(txtNroOrdenPago.Text) > 0 Then
       Set oRsTmp3 = mo_AdminFacturacion.ServiciosSeleccionarPorFiltro("idServicio=" & mo_cmbIdServicio.BoundText, sghPorCodigo)
       If oRsTmp3.Fields!IdServicio = 75 Then
          'solo Para PLANIFICACION FAMILIAR
          lcCola = ms_NroCola
       Else
          If (lbYaSeTransfirioHCdeUnServicioAotro = True) Then
              lcCola = ms_NroCola & "  PQTE (Ya se transfirió la Historia)"
          Else
              lcCola = ms_NroCola & Space(5) & "N°Ord.Pago: " & txtNroOrdenPago.Text
          End If
       End If
       oRsTmp3.Close
    Else
       lcCola = ms_NroCola & IIf(txtNroOrdenPago.Text = lcPagoCita And mo_Atenciones.IdFormaPago = 1, " " & txtNroOrdenPago.Text, "")
    End If
    If ucMensajeParpadeando2.Visible = True Then
       lcCola = lcCola & " (Cita Adicional)"
    End If
    
    Dim sNombreServicio As String
    Dim arrayServicio As Variant
    arrayServicio = Split(cmbIdServicio.Text, "=")
    sNombreServicio = Trim(arrayServicio(1))
    If lbImpresionCuenta = False And Val(lcBuscaParametro.SeleccionaFilaParametro(208)) = 1910 Then    'sullana
          oReporte.ImpresionFUAformato txtFechaIngreso.Text, txtHoraInicio.Text, lcPaciente, mo_Pacientes.NroHistoriaClinica, _
                                sNombreServicio, lcMedico, "CONSULTORIO EXTERNO", mo_Atenciones.idAtencion, _
                                txtNroOrdenPago.Text, mo_Atenciones.idCuentaAtencion, _
                                cmbFuenteFinanciamiento.Text & IIf(ucEPS1.Visible, mo_ReporteUtil.DevuelveEPScubre(ucEPS1.Porcentaje), ""), _
                                lcCola, ml_idUsuario, cmbIdTipoConsulta.Text, mo_Pacientes.FichaFamiliar, _
                                mo_Pacientes.idTipoNumeracion, wxParametro216, wxParametro306, _
                                mo_Pacientes.ApellidoPaterno, mo_Pacientes.ApellidoMaterno, mo_Pacientes.PrimerNombre, _
                                mo_Pacientes.SegundoNombre, mo_Pacientes.idTipoSexo, mo_Pacientes.FechaNacimiento, _
                                mo_Pacientes.nrodocumento, ml_IdMedico
    Else
       oReporte.ImpresionPreCuenta txtFechaIngreso.Text, txtHoraInicio.Text, lcPaciente, mo_Pacientes.NroHistoriaClinica, _
                                sNombreServicio, lcMedico, "CONSULTORIO EXTERNO", mo_Atenciones.idAtencion, _
                                txtNroOrdenPago.Text, mo_Atenciones.idCuentaAtencion, _
                                cmbFuenteFinanciamiento.Text & IIf(ucEPS1.Visible, mo_ReporteUtil.DevuelveEPScubre(ucEPS1.Porcentaje), ""), _
                                lcCola, ml_idUsuario, cmbIdTipoConsulta.Text, mo_Pacientes.FichaFamiliar, _
                                mo_Pacientes.idTipoNumeracion, wxParametro216, wxParametro306, False, ml_IdMedico
    End If
    Set oReporte = Nothing
    Set oRsTmp3 = Nothing
    If Not (chkPacienteNuevo.Value = 1 And mi_Opcion = sghAgregar) Then
       Me.Visible = False
    End If
    LimpiarVariablesDeMemoria
End Sub

Sub LimpiarVariablesDeMemoria()

End Sub

Private Sub txtNboleta_LostFocus()
    If Trim(txtNserie.Text) <> "" And Trim(txtNboleta.Text) <> "" Then
        Dim rsBuscaBoleta As New Recordset
        Dim oRsTmp As New Recordset
        Dim oRsTmp1 As New Recordset
        Dim lnIdEspecialidadServicio As Long
        Dim lcSql As String
        Dim lnIdProducto As Long
        Dim lcMensaje As String, lcCuentas As String
        lcMensaje = ""
        lcCuentas = ""
        Select Case mi_Opcion
        Case sghAgregar
        Case sghModificar
            If Trim(txtNserie.Text) = lcNserieAnt And Trim(txtNboleta.Text) = lcNboletaAnt Then
               lcMensaje = "xx"
            End If
        End Select
        If lcMensaje = "" Then
            mo_DOFacturacionPaquetes.IdComprobantePago = 0
            mo_DOFacturacionPaquetes.IdOrdenPago = 0
            mo_DOFacturacionPaquetes.idProducto = 0
            Set rsBuscaBoleta = mo_AdminCaja.CajaComprobantePagoServiciosPorNroSerieNroDocumento(txtNserie.Text, Trim(txtNboleta.Text))
            If rsBuscaBoleta.RecordCount > 0 Then
               If rsBuscaBoleta.Fields!idEstadoComprobante <> sghEstadosComprobante.sighEstadosComprobantePagado Then
                  lcMensaje = "Esa Boleta está ANULADA" & Chr(13)
               Else
                  lnIdEspecialidadServicio = 0
                  Set oRsTmp = mo_AdminServiciosComunes.ServiciosSeleccionarXidentificador(Val(mo_cmbIdServicio.BoundText))
                  If oRsTmp.RecordCount > 0 Then
                     lnIdEspecialidadServicio = oRsTmp.Fields!IdEspecialidad
                  End If
                  oRsTmp.Close
                  '
                  lnIdProducto = Val(mo_cmbIdTipoConsulta.BoundText)
                  '
                  Set oRsTmp = mo_AdminFacturacion.FacturacionPaquetesSeleccionarPorFiltro("idComprobantePago=" & Trim(Str(rsBuscaBoleta.Fields!IdComprobantePago)))
                  If oRsTmp.RecordCount = 0 Then
                     lcMensaje = "Esa Boleta no es un PAQUETE" & Chr(13)
                  Else
                     oRsTmp.Filter = "idEspecialidadServicio=" & lnIdEspecialidadServicio
                     If oRsTmp.RecordCount = 0 Then
                        lcMensaje = "Esa Boleta es un PAQUETE, pero la Especialidad del SERVICIO DE CITA es diferente a la de la Boleta" & Chr(13)
                     Else
                        oRsTmp.MoveFirst
                        Do While Not oRsTmp.EOF
                           If oRsTmp.Fields!AtencionId > 0 Then
                               If oRsTmp.Fields!idPuntoCarga = 6 Then
                                  lcCuentas = lcCuentas & oRsTmp.Fields!AtencionId & " , "
                               End If
                           Else
                                If oRsTmp.Fields!idPuntoCarga = 6 Then
                                    Set oRsTmp1 = mo_AdminCaja.CajaComprobantePagoSeleccionarPaquetes(oRsTmp.Fields!IdComprobantePago)
                                    If IsNull(oRsTmp1.Fields!idPaciente) Then
                                    ElseIf oRsTmp1.Fields!idPaciente <> ml_IdPaciente And oRsTmp1.Fields!idPaciente > 0 And ml_IdPaciente > 0 Then
                                        lcMensaje = "El Paciente de la Boleta (N° Historia " & oRsTmp1.Fields!NroHistoriaClinica & " " & Trim(oRsTmp1.Fields!ApellidoPaterno) & " " & Trim(oRsTmp1.Fields!ApellidoMaterno) & " " & Trim(oRsTmp1.Fields!PrimerNombre) & ")" & Chr(13) & "es diferente al de la CITA" & Chr(13)
                                        If MsgBox(lcMensaje + Chr(13) + Chr(13) + "¿ Continua ?", vbYesNo + vbQuestion, Me.Caption) = vbNo Then
                                           Exit Do
                                        Else
                                           lcMensaje = ""
                                        End If
                                    End If
                                    oRsTmp1.Close
                                    mo_DOFacturacionPaquetes.IdComprobantePago = oRsTmp.Fields!IdComprobantePago
                                    mo_DOFacturacionPaquetes.IdOrdenPago = oRsTmp.Fields!IdOrdenPago
                                    mo_DOFacturacionPaquetes.idProducto = oRsTmp.Fields!idProducto
                                    Exit Do
                                End If
                           End If
                           oRsTmp.MoveNext
                        Loop
                        If mo_DOFacturacionPaquetes.IdComprobantePago = 0 And lcMensaje = "" Then
                            lcMensaje = "Esa Boleta es un PAQUETE, pero ya se registró la CITA anteriormente" & Chr(13) & "chequee N° Cuentas: " & lcCuentas & Chr(13) & Chr(10)
                        End If
                     End If
                  End If
                  oRsTmp.Close
                End If
            Else
                lcMensaje = "No existe esa Boleta"
            End If
        End If
        Set oRsTmp = Nothing
        Set oRsTmp1 = Nothing
        If lcMensaje <> "" Then
            MsgBox lcMensaje, vbInformation, Me.Caption
            txtNboleta.Text = ""
            txtNboleta.SetFocus
        Else
            btnAceptar.SetFocus
        End If
    Else
        If mo_DOFacturacionPaquetes.IdComprobantePago > 0 Then
            mo_DOFacturacionPaquetes.IdComprobantePago = 0
            mo_DOFacturacionPaquetes.IdOrdenPago = 0
            mo_DOFacturacionPaquetes.idProducto = 0
        End If
    End If
End Sub

Function DevuelveNombreMedicoPlanilla(lnIdMedico As Long) As String
    Dim lcSql As String
    Dim oRsTmp As New Recordset
    Set oRsTmp = mo_ReglasDeProgMedica.MedicosSeleccionarPorIdMedicoPlanilla(lnIdMedico)
    DevuelveNombreMedicoPlanilla = ""
    If oRsTmp.RecordCount > 0 Then
       DevuelveNombreMedicoPlanilla = Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & " " & Trim(oRsTmp.Fields!Nombres) & " (" & Trim(oRsTmp.Fields!CodigoPlanilla) & ")"
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
End Function

'debb-Jamo
Function GrabaAtencionJamo() As Boolean
    If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE Then
        Dim oRsTmpBuscaAtencion As New Recordset
        Select Case mi_Opcion
        Case sghEliminar
             mo_DOAtencionesCE.idAtencion = mo_Atenciones.idAtencion
             mo_DOAtencionesCE.IdUsuarioAuditoria = mo_Atenciones.IdUsuarioAuditoria
             GrabaAtencionJamo = mo_AdminAdmision.AtencionCEeliminar(mo_DOAtencionesCE, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "IdAtencion: " & Trim(Str(ml_idAtencion)) & "(desde Atención)")
        End Select
        If GrabaAtencionJamo = False Then
           MsgBox "Falló al Grabar DATOS JAMO" & Chr(13) & mo_AdminAdmision.MensajeError
        End If
        Set oRsTmpBuscaAtencion = Nothing
    End If
End Function


Function GrabaAtencionPerinatal() As Boolean
        Dim oDoPerinatalAtencion As New DoPerinatalAtencion
        Select Case mi_Opcion
        Case sghEliminar
            oDoPerinatalAtencion.idPaciente = mo_Atenciones.idPaciente
            GrabaAtencionPerinatal = mo_AdminAdmision.PerinatalCEeliminar(oDoPerinatalAtencion, mo_Atenciones.idAtencion)
        End Select
        If GrabaAtencionPerinatal = False Then
           MsgBox "Falló al Grabar PERINATAL" & Chr(13) & mo_AdminAdmision.MensajeError
        End If
        Set oDoPerinatalAtencion = Nothing
End Function



Private Sub UcPacientesSunasa1_SePresionoTeclaEspecial(KeyCode As Integer)
   AdministrarKeyPreview KeyCode
End Sub


Function DevuelveNroRecetasGeneradas() As String
    DevuelveNroRecetasGeneradas = ""
    If lnRecetaRayosX > 0 Then
       DevuelveNroRecetasGeneradas = DevuelveNroRecetasGeneradas & Chr(13) & "N° Receta para: Rayos X: " & Trim(Str(lnRecetaRayosX))
    End If
    If lnRecetaEcografiaO > 0 Then
       DevuelveNroRecetasGeneradas = DevuelveNroRecetasGeneradas & Chr(13) & "N° Receta para: Ecografía Obstétrica: " & Trim(Str(lnRecetaEcografiaO))
    End If
    If lnRecetaEcografiaG > 0 Then
       DevuelveNroRecetasGeneradas = DevuelveNroRecetasGeneradas & Chr(13) & "N° Receta para: Ecografía General: " & Trim(Str(lnRecetaEcografiaG))
    End If
    If lnRecetaTomografia > 0 Then
       DevuelveNroRecetasGeneradas = DevuelveNroRecetasGeneradas & Chr(13) & "N° Receta para: Tomografía: " & Trim(Str(lnRecetaTomografia))
    End If
    If lnRecetaAnatomiaP > 0 Then
       DevuelveNroRecetasGeneradas = DevuelveNroRecetasGeneradas & Chr(13) & "N° Receta para: Anatomía Patológica: " & Trim(Str(lnRecetaAnatomiaP))
    End If
    If lnRecetaPatologiaC > 0 Then
       DevuelveNroRecetasGeneradas = DevuelveNroRecetasGeneradas & Chr(13) & "N° Receta para: Patológia Clínica: " & Trim(Str(lnRecetaPatologiaC))
    End If
    If lnRecetaBancoS > 0 Then
       DevuelveNroRecetasGeneradas = DevuelveNroRecetasGeneradas & Chr(13) & "N° Receta para: Banco de Sangre: " & Trim(Str(lnRecetaBancoS))
    End If
    If lnRecetaFarmacia > 0 Then
       DevuelveNroRecetasGeneradas = DevuelveNroRecetasGeneradas & Chr(13) & "N° Receta para: Farmacia: " & Trim(Str(lnRecetaFarmacia))
    End If
End Function

Sub InicilizarParametros()
    wxParametro205 = lcBuscaParametro.SeleccionaFilaParametro(205)
    wxParametro208 = lcBuscaParametro.SeleccionaFilaParametro(208)
    wxParametro211 = lcBuscaParametro.SeleccionaFilaParametro(211)
    wxParametro216 = lcBuscaParametro.SeleccionaFilaParametro(216)
    wxParametro237 = lcBuscaParametro.SeleccionaFilaParametro(237)
    wxParametro242 = lcBuscaParametro.SeleccionaFilaParametro(242)
    wxParametro258 = lcBuscaParametro.SeleccionaFilaParametro(258)
    wxParametro259 = lcBuscaParametro.SeleccionaFilaParametro(259)
    wxParametro274 = lcBuscaParametro.SeleccionaFilaParametro(274)
    wxParametro275 = lcBuscaParametro.SeleccionaFilaParametro(275)
    wxParametro276 = lcBuscaParametro.SeleccionaFilaParametro(276)
    wxParametro282 = lcBuscaParametro.SeleccionaFilaParametro(282)
    wxParametro287 = lcBuscaParametro.SeleccionaFilaParametro(287)
    wxParametro296 = lcBuscaParametro.SeleccionaFilaParametro(296)
    wxParametro302 = lcBuscaParametro.SeleccionaFilaParametro(302)
    wxParametro306 = lcBuscaParametro.SeleccionaFilaParametro(306)
    wxParametro312 = lcBuscaParametro.SeleccionaFilaParametro(312)
    wxParametro322 = lcBuscaParametro.SeleccionaFilaParametro(322)
    wxParametro323 = lcBuscaParametro.SeleccionaFilaParametro(323)
    wxParametro333 = lcBuscaParametro.SeleccionaFilaParametro(333)
    wxParametro336 = lcBuscaParametro.SeleccionaFilaParametro(336)
    wxParametro351 = lcBuscaParametro.SeleccionaFilaParametro(351)
    wxParametro358 = lcBuscaParametro.SeleccionaFilaParametro(358)
    wxParametro359 = lcBuscaParametro.SeleccionaFilaParametro(359)
    wxParametro517 = lcBuscaParametro.SeleccionaFilaParametro(517)  'franklin 2017
    wxParametro526 = lcBuscaParametro.SeleccionaFilaParametro(526)
    wxParametro539 = lcBuscaParametro.SeleccionaFilaParametro(539)
    wxParametro540 = lcBuscaParametro.SeleccionaFilaParametro(540)
    wxParametro580 = lcBuscaParametro.SeleccionaFilaParametro(580)
    wxParametroJAMO = lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    ldFechaActualServidor = lcBuscaParametro.RetornaFechaServidorSQL
    'mgaray
    wxParametroBusqRapida = UCase(lcBuscaParametro.SeleccionaFilaParametro(344))
    '
End Sub

Private Sub UcSISafiliacion1_OnLostFocus(lcDisa As String, lcLote As String, lcNumero As String)
   lnAfiliacionSIS1 = lcDisa
   lnAfiliacionSIS2 = lcLote
   lnAfiliacionSIS3 = lcNumero
   Me.chkBuscarEnSIS.Value = 1
'   If lnAfiliacionSIS3 <> "" Then
'      btnBuscarPaciente_Click
'      On Error Resume Next
'      Me.grdPacientesEncontrados.SetFocus
'   End If
End Sub

Sub HaceVisibleOnoBotonFUA()
        btnImprimeFichaSIS.Visible = False
        ucSISfuaCodPrestacion1.Visible = False
        If Val(mo_cmbIdFuentesFinanciamiento.BoundText) = sghFuenteFinanciamiento.sghFFSIS Then
           If lbElMedicoNOregistraFUA = "S" And mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroCitaCE Then
              btnImprimeFichaSIS.Visible = True
              'If mi_Opcion = sghModificar Then
                  ucSISfuaCodPrestacion1.Visible = True
              'End If
           ElseIf mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE Then
              btnImprimeFichaSIS.Visible = True
           End If
        End If
        If ucSISfuaCodPrestacion1.Visible = True Then
           Dim lcSexo As String, ml_edad_En_YYYYMMDD As String
           If sighEntidades.EsFecha(Me.ucPacientesDetalle1.DevuelveFechaNacimiento, "DD/MM/AAAA") = True Then
              ml_edad_En_YYYYMMDD = sighEntidades.EdadActualEnFormatoYYYYMMDD(CDate(Format(Me.ucPacientesDetalle1.DevuelveFechaNacimiento, "dd/mm/yyyy hh:mm")), CDate(Format(txtFechaIngreso.Text & " " & txtHoraInicio.Text, "dd/mm/yyyy hh:mm")))
           End If
           lcSexo = IIf(Left(Me.ucPacientesDetalle1.DevuelveSexo, 1) = 1, "M", "F")
           Me.ucSISfuaCodPrestacion1.ReglasDeConsistenciasAntesDeCargarFormulario ml_TipoServicio, lcSexo, ml_edad_En_YYYYMMDD
           If mi_Opcion <> sghAgregar And mo_DoAtencionDatosAdicionales.FuaCodigoPrestacion <> "" Then
              ucSISfuaCodPrestacion1.CodigoPrestacion = mo_DoAtencionDatosAdicionales.FuaCodigoPrestacion
           End If
        End If
End Sub

Function EpisodioClinicoDevuelveDatos() As EpisodioClinico
        Dim oEpisodioClinico As EpisodioClinico
        oEpisodioClinico.idEpisodio = Me.UcEpisodioClinico1.idEpisodio
        oEpisodioClinico.lbCierreEpisodio = Me.UcEpisodioClinico1.lbCierreEpisodio
        oEpisodioClinico.lbNuevoEpisodio = Me.UcEpisodioClinico1.lbNuevoEpisodio
        EpisodioClinicoDevuelveDatos = oEpisodioClinico
End Function

'mgaray201503
Public Function UsuarioActualEsCajero() As Boolean
    UsuarioActualEsCajero = Principal.UsuarioActualEsCajero
End Function


'franklin 2017
Sub BuscaMedicoRerencia(lcIdColegio As String)
    If Len(txtMedicoRef.Text) >= 1 Then
        Dim oRsTmp112 As New Recordset
        Dim lnId As Integer, lnIdex As Integer
        Set oRsTmp112 = mo_ReglasSISgalenhos.a_resatencionSeleccionarPorColegiatura(txtMedicoRef.Text)
        cmbMedicoRef.Clear
        lnId = 1: lnIdex = 0
        If oRsTmp112.RecordCount > 0 Then
           oRsTmp112.MoveFirst
           Do While Not oRsTmp112.EOF
              If Val(oRsTmp112!pers_IdTipoPersonalSalud) = Val(lcIdColegio) Then
                 lnIdex = lnId
              End If
              cmbMedicoRef.AddItem oRsTmp112!Medico
              oRsTmp112.MoveNext
              lnId = lnId + 1
           Loop
           If oRsTmp112.RecordCount = 1 Then
              cmbMedicoRef.ListIndex = 0
           ElseIf lnIdex > 0 Then
              cmbMedicoRef.ListIndex = lnIdex - 1
           End If
        End If
        oRsTmp112.Close
        Set oRsTmp112 = Nothing
    End If
End Sub
'franklin 2017
Sub ActualizaCitas_Atencion(oConexion As Connection)
    If wxParametro517 = "" Then
       Exit Sub
    End If
    On Error GoTo errActCita
    Dim oConexionExt As New Connection
    Dim oRsTmp65 As New Recordset
    Dim oDoMedico As New DOMedico
    Dim oMedicos As New Medicos
    Dim oDOEmpleado As New dOEmpleado
    Dim oEmpleados As New Empleados
    Dim lcDNImedico As String
    Dim lcSql As String
    Dim lcErrorSql As String
    Dim lbEsNuevo As Boolean
    lcErrorSql = "inicio"
    lcDNImedico = ""
    Set oEmpleados.Conexion = oConexion
    Set oMedicos.Conexion = oConexion
    oDoMedico.idMedico = mo_Atenciones.IdMedicoIngreso
    oDoMedico.IdUsuarioAuditoria = mo_Atenciones.IdUsuarioAuditoria
    If oMedicos.SeleccionarPorId(oDoMedico) Then
       oDOEmpleado.IdEmpleado = oDoMedico.IdEmpleado
       oDOEmpleado.IdUsuarioAuditoria = mo_Atenciones.IdUsuarioAuditoria
       If oEmpleados.SeleccionarPorId(oDOEmpleado) Then
          lcDNImedico = Trim(oDOEmpleado.DNI)
       End If
    End If
    lcErrorSql = "busco dni medico"
    oConexionExt.CommandTimeout = 900
    oConexionExt.Open wxParametro517
    lcErrorSql = "paso abrir conexion"
    If mi_Opcion = sghModificar Or mi_Opcion = sghEliminar Then
        lcSql = "select * from citas_atencion where horaCita='" & mo_Atenciones.HoraIngreso & "'" & _
          " and servicioId=" & mo_Atenciones.IdServicioIngreso & _
          " and historia=" & mo_Pacientes.NroHistoriaClinica
        oRsTmp65.Open lcSql, oConexionExt, adOpenKeyset, adLockOptimistic
        lcErrorSql = "paso abrir Modificar/eliminar"
        oRsTmp65.Filter = "fechaCita='" & mo_Atenciones.FechaIngreso & "'"
        lcErrorSql = "paso filtro por fecha"
        lbEsNuevo = False
        If oRsTmp65.RecordCount = 0 Then
           lbEsNuevo = True
        End If
    Else
        lcSql = "select * from citas_atencion"
        oRsTmp65.Open lcSql, oConexionExt, adOpenKeyset, adLockOptimistic
        lcErrorSql = "paso abrir toda la tabla"
        lbEsNuevo = True
    End If
    If mi_Opcion = sghAgregar Or mi_Opcion = sghModificar Then
         If lbEsNuevo = True Then
            oRsTmp65.AddNew
            oRsTmp65!fechaCita = mo_Atenciones.FechaIngreso
            oRsTmp65!horaCita = mo_Atenciones.HoraIngreso
            oRsTmp65!ServicioId = mo_Atenciones.IdServicioIngreso
            oRsTmp65!Servicio = Trim(Mid(cmbIdServicio.Text, InStr(cmbIdServicio.Text, "=") + 1, 100))
            oRsTmp65!Historia = mo_Pacientes.NroHistoriaClinica
            oRsTmp65!EstadoCita = "Citado"
            oRsTmp65!DxCodigo = ""
            oRsTmp65!Dx = ""
         End If
         oRsTmp65!MedicoDni = lcDNImedico
         If mo_Pacientes.IdDocIdentidad = 1 Then
            oRsTmp65!DNI = Left(mo_Pacientes.nrodocumento, 8)
         Else
            oRsTmp65!DNI = ""
         End If
         oRsTmp65!TipoSeguro = cmbFuenteFinanciamiento.Text
         oRsTmp65!TipoSeguroID = mo_Atenciones.IdFuenteFinanciamiento
         oRsTmp65!ApellidoPaterno = mo_Pacientes.ApellidoPaterno
         oRsTmp65!ApellidoMaterno = mo_Pacientes.ApellidoMaterno
         oRsTmp65!Nombres = Left(mo_Pacientes.PrimerNombre & " " & mo_Pacientes.SegundoNombre, 60)
         oRsTmp65!FechaNacimiento = mo_Pacientes.FechaNacimiento
         oRsTmp65!Sexo = IIf(mo_Pacientes.idTipoSexo = 1, "Masculino", "Femenino")
         oRsTmp65!Direccion = Left(mo_Pacientes.DireccionDomicilio, 100)
         oRsTmp65!DireccionUbigeo = mo_Pacientes.IdDistritoDomicilio
         oRsTmp65!cuenta = mo_Atenciones.idCuentaAtencion
         oRsTmp65.Update
    Else
         oRsTmp65.Delete
         oRsTmp65.Update
    End If
    lcErrorSql = "paso grabar"
    oRsTmp65.Close
    oConexionExt.Close
    Set oConexionExt = Nothing
    Set oRsTmp65 = Nothing
    lcErrorSql = "cerro todos las variables"
    Exit Sub
errActCita:
    'MsgBox lcErrorSql & Chr(13) & Err.Description
    lblNroAtencion.Caption = lcErrorSql & Chr(13) & Err.Description
End Sub

'SCCQ 28/04/2021 Cambio 64 Inicio
Function CargarProcedimientoFUA(idProductoFUA As String) As DOFacturacionServicios
Dim oServicio As New DOFacturacionServicios
Dim oRsBuscaSeguro As New ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim PrSeguro As Double
            
            With oServicio
                .idAtencion = Me.idAtencion
                .IdFacturacionServicio = 0
                .IdFuenteFinanciamiento = Val(mo_cmbIdFuentesFinanciamiento.BoundText)
                .idTipoFinanciamiento = Val(mo_cmbIdFormaPago.BoundText)
                .Cantidad = 1
                .idProducto = idProductoFUA
                .IdUsuarioAuditoria = ml_idUsuario
                .idestadofacturacion = sghEstadoFacturacion.sghRegistrado
                .FechaAutorizaPendiente = 0
                .FechaAutorizaSeguro = 0
                .IdCentroCosto = 0
                .IdUsuarioAutorizaPendiente = 0
                .IdUsuarioAutorizaSeguro = 0
                .PrecioUnitario = 0
                .TotalPorPagar = 0
                .idPuntoCarga = 6
                '********pone Seguros en forma automatica, sin necesidad de ir a SEGUROS-inicio
                If mi_Opcion = sghModificar Then
                   .IdOrden = lnIdFactServicios
                End If
                PrSeguro = 0
                oConexion.Open sighEntidades.CadenaConexion
                oConexion.CursorLocation = adUseClient
                If Val(mo_cmbIdFormaPago.BoundText) = 9 Then
                   'Si es EXONERACIONES tomará el PRECIO de un Paciente Normal
                   Set oRsBuscaSeguro = mo_AdminFacturacion.CatalogoServiciosHospSeleccionarXidProductoIdTipoFinanciamiento(.idProducto, 1, oConexion)
                Else
                   Set oRsBuscaSeguro = mo_AdminFacturacion.CatalogoServiciosHospSeleccionarXidProductoIdTipoFinanciamiento(.idProducto, Val(mo_cmbIdFormaPago.BoundText), oConexion)
                End If
                If oRsBuscaSeguro.RecordCount > 0 Then
                   PrSeguro = oRsBuscaSeguro.Fields!PrecioUnitario
                End If
                oRsBuscaSeguro.Close
                Set oRsBuscaSeguro = Nothing
                oConexion.Close
                Set oConexion = Nothing
                .idTipoFinanciamiento = Val(mo_cmbIdFormaPago.BoundText)
                .CantidadSIS = 0
                .precioSIS = 0
                .ImporteSIS = 0
                .CantidadSOAT = 0
                .PrecioSOAT = 0
                .ImporteSOAT = 0
                .importeEXO = 0
                .cantidadConv = 0
                .precConv = 0
                .ImporteConv = 0
                Select Case Val(mo_cmbIdFormaPago.BoundText)
                Case 1  'Contado
                Case 2  'SIS
                     If PrSeguro > 0 Then
                        .CantidadSIS = 1
                        .precioSIS = PrSeguro
                        .ImporteSIS = PrSeguro
                        .idestadofacturacion = 10
                        .FechaAutorizaSeguro = Now
                     Else
                        .idTipoFinanciamiento = 1
                        'mo_Atenciones.IdFormaPago = 1
                        'mo_cmbIdFormaPago.BoundText = "1"
                     End If
             
                Case 9  'Exonerados
                    .importeEXO = PrSeguro
                    .idestadofacturacion = 10
                    .idTipoFinanciamiento = 1
                    .FechaAutorizaEXO = Now
                End Select
                '********pone Seguros en forma automatica, sin necesidad de ir a SEGUROS-fin
            End With
            
            Set CargarProcedimientoFUA = oServicio

End Function
'SCCQ 28/04/2021 Cambio 64 Fin

