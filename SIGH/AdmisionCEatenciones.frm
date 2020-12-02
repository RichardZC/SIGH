VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{22ACD161-99EB-11D2-9BB3-00400561D975}#1.0#0"; "PVCALE~1.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form AdmisionCEatenciones 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9585
   ClientLeft      =   1830
   ClientTop       =   -105
   ClientWidth     =   15975
   ControlBox      =   0   'False
   Icon            =   "AdmisionCEatenciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   15975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SISGalenPlus.ucPacientesCtasPDF ucPacientesCtasPDF1 
      Height          =   1755
      Left            =   12360
      TabIndex        =   92
      Top             =   4965
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   3096
   End
   Begin VB.TextBox lblMedico 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   12360
      MaxLength       =   1000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   91
      Top             =   6795
      Width           =   3600
   End
   Begin VB.Frame fraTriaje 
      Caption         =   "Triaje"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3510
      Left            =   12360
      TabIndex        =   6
      Top             =   0
      Width           =   3570
      Begin SISGalenPlus.ucTriajeVisor ucTriajeVisorCE 
         Height          =   3285
         Left            =   75
         TabIndex        =   42
         Top             =   180
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   5794
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2190
      Left            =   12360
      TabIndex        =   2
      Top             =   7395
      Width           =   3615
      Begin VB.CommandButton btnAntecedentesPersonales 
         Caption         =   "Anteced."
         DisabledPicture =   "AdmisionCEatenciones.frx":0CCA
         DownPicture     =   "AdmisionCEatenciones.frx":112A
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
         Left            =   1530
         Picture         =   "AdmisionCEatenciones.frx":159F
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Antecedentes Niño"
         Top             =   1530
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.CommandButton cmdGenerarPlanAtencion 
         Caption         =   "P. Aten."
         DisabledPicture =   "AdmisionCEatenciones.frx":1FA1
         DownPicture     =   "AdmisionCEatenciones.frx":2401
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
         Left            =   2655
         Picture         =   "AdmisionCEatenciones.frx":2B03
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Generar/Actualizar Plan Atención"
         Top             =   1530
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.CommandButton btnImprimeFichaSIS 
         Caption         =   "Imp. FUA"
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
         Left            =   2295
         Picture         =   "AdmisionCEatenciones.frx":3505
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   855
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.CommandButton btnImprimeAtencion 
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
         Height          =   615
         Left            =   90
         Picture         =   "AdmisionCEatenciones.frx":39DE
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   855
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
         Left            =   90
         Picture         =   "AdmisionCEatenciones.frx":3EB7
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1530
         Width           =   1245
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar "
         DisabledPicture =   "AdmisionCEatenciones.frx":4441
         DownPicture     =   "AdmisionCEatenciones.frx":4905
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
         Left            =   2295
         Picture         =   "AdmisionCEatenciones.frx":4DF1
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   150
         Width           =   1245
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "AdmisionCEatenciones.frx":52DD
         DownPicture     =   "AdmisionCEatenciones.frx":573D
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
         Left            =   90
         Picture         =   "AdmisionCEatenciones.frx":5BB2
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   150
         Width           =   1245
      End
   End
   Begin TabDlg.SSTab TabAtencion 
      Height          =   9600
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   12345
      _ExtentX        =   21775
      _ExtentY        =   16933
      _Version        =   393216
      Tabs            =   5
      Tab             =   2
      TabsPerRow      =   5
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
      TabCaption(0)   =   "3.1 Anam/Ex.Físico"
      TabPicture(0)   =   "AdmisionCEatenciones.frx":6027
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame17"
      Tab(0).Control(1)=   "Frame16"
      Tab(0).Control(2)=   "Frame7"
      Tab(0).Control(3)=   "Frame6"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "3.2 Diagnósticos"
      TabPicture(1)   =   "AdmisionCEatenciones.frx":6043
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "btnguardarcs"
      Tab(1).Control(1)=   "grdServicios"
      Tab(1).Control(2)=   "TabDx"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "3.3 Ordenes Médicas"
      TabPicture(2)   =   "AdmisionCEatenciones.frx":605F
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "txtCitaExClinicos"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "btnImprimir"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "btnImprimirOrden"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "UcRecetas1"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "btnAgregaApoyoDx"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "3.4 Tratamiento"
      TabPicture(3)   =   "AdmisionCEatenciones.frx":607B
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame10"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "3.5 Destino Atención"
      TabPicture(4)   =   "AdmisionCEatenciones.frx":6097
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblNroAtencion"
      Tab(4).Control(1)=   "ucCitasLista1"
      Tab(4).Control(2)=   "UcEpisodioClinico1"
      Tab(4).Control(3)=   "Frame1"
      Tab(4).Control(4)=   "fraDatosReferenciaDestino"
      Tab(4).Control(5)=   "Frame11"
      Tab(4).ControlCount=   6
      Begin VB.CommandButton btnguardarcs 
         Caption         =   "Guardar "
         Height          =   360
         Left            =   -63360
         TabIndex        =   95
         Top             =   9120
         Visible         =   0   'False
         Width           =   510
      End
      Begin SISGalenPlus.ucCatalogos grdServicios 
         Height          =   3015
         Left            =   -75000
         TabIndex        =   93
         Top             =   6600
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   5318
      End
      Begin VB.CommandButton btnAgregaApoyoDx 
         DisabledPicture =   "AdmisionCEatenciones.frx":60B3
         DownPicture     =   "AdmisionCEatenciones.frx":649C
         Height          =   390
         Left            =   45
         Picture         =   "AdmisionCEatenciones.frx":68A8
         Style           =   1  'Graphical
         TabIndex        =   90
         ToolTipText     =   "Agrega RECETA"
         Top             =   450
         Width           =   300
      End
      Begin VB.Frame Frame11 
         Caption         =   "Otras observaciones"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -74895
         TabIndex        =   74
         Top             =   5250
         Width           =   11655
         Begin VB.TextBox txtCitaObservaciones 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2205
            Left            =   120
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   75
            Top             =   255
            Width           =   11415
         End
      End
      Begin VB.Frame fraDatosReferenciaDestino 
         Caption         =   " Destino de  referencia "
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2340
         Left            =   -74880
         TabIndex        =   61
         Top             =   1725
         Width           =   5385
         Begin VB.CommandButton btnBuscarEstablecimientoDestino 
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
            Left            =   2325
            Picture         =   "AdmisionCEatenciones.frx":6CB4
            Style           =   1  'Graphical
            TabIndex        =   89
            Top             =   690
            Width           =   330
         End
         Begin VB.CommandButton cmdImpresionReferencias 
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
            Left            =   75
            Picture         =   "AdmisionCEatenciones.frx":723E
            Style           =   1  'Graphical
            TabIndex        =   87
            ToolTipText     =   "Imprimir recetas de farmacia"
            Top             =   1785
            Width           =   405
         End
         Begin VB.ComboBox cmbIdTipoReferenciaDestino 
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
            TabIndex        =   68
            Top             =   300
            Width           =   1695
         End
         Begin VB.TextBox txtIdEstablecimientoDestino 
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
            TabIndex        =   67
            Top             =   690
            Width           =   675
         End
         Begin VB.TextBox txtNombreDestinoReferencia 
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
            Left            =   2670
            TabIndex        =   66
            Top             =   690
            Width           =   2625
         End
         Begin VB.TextBox txtNroReferenciaDestino 
            Height          =   315
            Left            =   4260
            MaxLength       =   20
            TabIndex        =   62
            Top             =   330
            Width           =   1020
         End
         Begin PVCOMBOLibCtl.PVComboBox cmbServicioReferenciaD 
            Height          =   330
            Left            =   1650
            TabIndex        =   63
            Top             =   1050
            Width           =   3690
            _Version        =   524288
            _cx             =   6509
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
         Begin MSMask.MaskEdBox txtFextension 
            Height          =   315
            Left            =   1650
            TabIndex        =   64
            Top             =   1425
            Width           =   1335
            _ExtentX        =   2355
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
         Begin MSMask.MaskEdBox txtFtramite 
            Height          =   315
            Left            =   3960
            TabIndex        =   65
            Top             =   1425
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   11
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
         Begin VB.Label lblServicioReferencia0 
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
            Left            =   150
            TabIndex        =   79
            Top             =   1080
            Width           =   1485
         End
         Begin VB.Label lblFextension0 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F. extensión"
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
            Left            =   150
            TabIndex        =   78
            Top             =   1410
            Width           =   1005
         End
         Begin VB.Label lblFtramite0 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F. Trámite"
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
            Left            =   3045
            TabIndex        =   77
            Top             =   1470
            Width           =   840
         End
         Begin VB.Label lblIdEstablecimientoDestino 
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
            Left            =   150
            TabIndex        =   71
            Top             =   720
            Width           =   1380
         End
         Begin VB.Label lblIdTipoReferenciaDestino 
            AutoSize        =   -1  'True
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
            Left            =   150
            TabIndex        =   70
            Top             =   360
            Width           =   1230
         End
         Begin VB.Label lblReferenciaO 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N° Refer"
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
            Left            =   3510
            TabIndex        =   69
            Top             =   360
            Width           =   705
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1305
         Left            =   -74880
         TabIndex        =   52
         Top             =   360
         Width           =   5325
         Begin VB.ComboBox cmbIdDestinoAtencion 
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
            Left            =   1380
            TabIndex        =   55
            Top             =   570
            Width           =   3870
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Caption         =   "Alta definitiva"
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
            TabIndex        =   54
            Top             =   330
            Width           =   1455
         End
         Begin VB.TextBox txtProximaCita 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   53
            Top             =   930
            Width           =   435
         End
         Begin PVATLCALENDARLib.PVCalendar Calendario 
            Height          =   4455
            Left            =   150
            TabIndex        =   56
            TabStop         =   0   'False
            ToolTipText     =   "Haga click para seleccionar el día que desea asignar la cita"
            Top             =   1260
            Width           =   5085
            _Version        =   524288
            BorderStyle     =   1
            Appearance      =   1
            FirstDay        =   1
            Frame           =   1
            SelectMode      =   2
            DisplayFormat   =   0
            DateOrientation =   0
            CustomTextOrientation=   2
            ImageOrientation=   8
            DOWText0        =   "Dom"
            DOWText1        =   "Lun"
            DOWText2        =   "Mar"
            DOWText3        =   "Mie"
            DOWText4        =   "Jue"
            DOWText5        =   "Vie"
            DOWText6        =   "Sab"
            MonthText0      =   "Enero"
            MonthText1      =   "Febrero"
            MonthText2      =   "MArzo"
            MonthText3      =   "Abril"
            MonthText4      =   "Mayo"
            MonthText5      =   "Junio"
            MonthText6      =   "Julio"
            MonthText7      =   "Agosto"
            MonthText8      =   "Setiembre"
            MonthText9      =   "Octubre"
            MonthText10     =   "Noviembre"
            MonthText11     =   "Diciembre"
            HeaderBackColor =   15780518
            HeaderForeColor =   0
            DisplayBackColor=   11888424
            DisplayForeColor=   0
            DayBackColor    =   16577517
            DayForeColor    =   0
            SelectedDayForeColor=   16777215
            SelectedDayBackColor=   16737792
            BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DOWFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DaysFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MultiLineText   =   -1  'True
            EditMode        =   0
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label46 
            Caption         =   "Destino"
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
            Left            =   180
            TabIndex        =   60
            Top             =   630
            Width           =   1005
         End
         Begin VB.Label Label32 
            Caption         =   "Próxima Cita"
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
            Left            =   180
            TabIndex        =   59
            Top             =   990
            Width           =   1005
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "días"
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
            Left            =   1860
            TabIndex        =   58
            Top             =   960
            Width           =   360
         End
         Begin VB.Label lblProximaCita 
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
            Left            =   4800
            TabIndex        =   57
            Top             =   960
            Width           =   360
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Indicaciones del Tratamiento"
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
         Left            =   -74880
         TabIndex        =   50
         Top             =   360
         Width           =   5775
         Begin VB.TextBox TxtCitaTratamiento 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4665
            Left            =   240
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   51
            Top             =   300
            Width           =   5415
         End
      End
      Begin SISGalenPlus.UcRecetaCE UcRecetas1 
         Height          =   9195
         Left            =   390
         TabIndex        =   49
         Top             =   315
         Width           =   11685
         _ExtentX        =   20611
         _ExtentY        =   16219
      End
      Begin VB.CommandButton btnImprimirOrden 
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
         Left            =   11400
         Picture         =   "AdmisionCEatenciones.frx":7717
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Imprimir ordenes médicas"
         Top             =   1800
         Width           =   440
      End
      Begin VB.CommandButton btnImprimir 
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
         Left            =   11350
         Picture         =   "AdmisionCEatenciones.frx":7BF0
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Imprimir recetas de farmacia"
         Top             =   600
         Width           =   440
      End
      Begin VB.Frame Frame6 
         Caption         =   "Motivo de la Consulta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2565
         Left            =   -74910
         TabIndex        =   38
         Top             =   2760
         Width           =   5985
         Begin VB.TextBox txtCitaMotivo 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2235
            Left            =   135
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   39
            Top             =   285
            Width           =   5745
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Exámen Clínico"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2565
         Left            =   -68880
         TabIndex        =   36
         Top             =   2760
         Width           =   5595
         Begin VB.TextBox txtCitaExamenClinico 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2235
            Left            =   150
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   37
            Top             =   270
            Width           =   5295
         End
      End
      Begin VB.Frame Frame16 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Antecedentes personales"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2385
         Left            =   -74880
         TabIndex        =   10
         Top             =   360
         Width           =   11595
         Begin VB.TextBox txtAntecedentes 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   6735
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Top             =   1440
            Width           =   4725
         End
         Begin VB.TextBox txtantecedQuirurgico 
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
            Left            =   1050
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   15
            Top             =   240
            Width           =   4845
         End
         Begin VB.TextBox txtantecedAlergico 
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
            Left            =   6750
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   240
            Width           =   4725
         End
         Begin VB.TextBox txtantecedPatologico 
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
            Left            =   1050
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Top             =   840
            Width           =   4845
         End
         Begin VB.TextBox txtantecedFamiliar 
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
            Left            =   6750
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   840
            Width           =   4725
         End
         Begin VB.TextBox txtantecedObstetrico 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   1050
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   1440
            Width           =   4845
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Otros"
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
            Left            =   6270
            TabIndex        =   22
            Top             =   1500
            Width           =   450
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quirúrgicos"
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
            TabIndex        =   21
            Top             =   270
            Width           =   900
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alergias "
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
            Left            =   6120
            TabIndex        =   20
            Top             =   300
            Width           =   675
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Patológicos"
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
            TabIndex        =   19
            Top             =   900
            Width           =   915
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Familiares"
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
            Left            =   6000
            TabIndex        =   18
            Top             =   900
            Width           =   750
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Obstétricos"
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
            TabIndex        =   17
            Top             =   1500
            Width           =   930
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Antecedentes relacionados a la Consulta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Left            =   -74910
         TabIndex        =   8
         Top             =   5400
         Width           =   11625
         Begin VB.TextBox txtCitaAntecedente 
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
            Left            =   150
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   270
            Width           =   11415
         End
      End
      Begin TabDlg.SSTab TabDx 
         Height          =   6255
         Left            =   -75000
         TabIndex        =   23
         Top             =   360
         Width           =   12225
         _ExtentX        =   21564
         _ExtentY        =   11033
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
         TabCaption(0)   =   "3.2.1 Información morbilidad"
         TabPicture(0)   =   "AdmisionCEatenciones.frx":80C9
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label34"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblHemoglobina"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "grdOtrosCpt"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Frame9"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "txtNroHijos"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "btnCpt"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Frame2"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "btnQuitarCpt"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "UcDiagnosticoDetalle1"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "txtHemoglobina"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Frame"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "cmdRegistraActividades"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Command1"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).ControlCount=   13
         TabCaption(1)   =   "3.2.2 Módulo Niño Sano"
         TabPicture(1)   =   "AdmisionCEatenciones.frx":80E5
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "ucPerinatal1"
         Tab(1).Control(1)=   "ucPerinatalAS1"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "3.2.3 Módulo Materno"
         TabPicture(2)   =   "AdmisionCEatenciones.frx":8101
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "UcProgramaMaterno"
         Tab(2).ControlCount=   1
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   360
            Left            =   13560
            TabIndex        =   94
            Top             =   7200
            Width           =   990
         End
         Begin VB.CommandButton cmdRegistraActividades 
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
            Left            =   11235
            Picture         =   "AdmisionCEatenciones.frx":811D
            Style           =   1  'Graphical
            TabIndex        =   88
            ToolTipText     =   "Llena ACTIVIDADES HIS desde cero"
            Top             =   3600
            Width           =   450
         End
         Begin VB.Frame Frame 
            Height          =   525
            Left            =   8820
            TabIndex        =   83
            Top             =   3480
            Width           =   2295
            Begin VB.TextBox txtEdad1 
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
               Left            =   120
               TabIndex        =   86
               Top             =   165
               Width           =   495
            End
            Begin VB.ComboBox cmbTipoEdad1 
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
               ItemData        =   "AdmisionCEatenciones.frx":86A7
               Left            =   660
               List            =   "AdmisionCEatenciones.frx":86B7
               TabIndex        =   85
               Text            =   "Combo"
               Top             =   150
               Width           =   1110
            End
            Begin VB.CommandButton cmdMuestraActividades 
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
               Left            =   1785
               Picture         =   "AdmisionCEatenciones.frx":86D5
               Style           =   1  'Graphical
               TabIndex        =   84
               ToolTipText     =   "Lista de ACTIVIDADES HIS de acuerdo a la EDAD, PESO  y UPS"
               Top             =   150
               Width           =   450
            End
         End
         Begin VB.TextBox txtHemoglobina 
            Alignment       =   1  'Right Justify
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
            Left            =   10245
            TabIndex        =   81
            Top             =   4350
            Visible         =   0   'False
            Width           =   1455
         End
         Begin SISGalenPlus.ucPerinatal ucPerinatal1 
            Height          =   3120
            Left            =   -74610
            TabIndex        =   33
            Top             =   2385
            Visible         =   0   'False
            Width           =   9090
            _ExtentX        =   20558
            _ExtentY        =   12515
         End
         Begin SISGalenPlus.ucPerinatalAS ucPerinatalAS1 
            Height          =   1440
            Left            =   -74745
            TabIndex        =   76
            Top             =   480
            Width           =   9750
            _ExtentX        =   17198
            _ExtentY        =   2540
         End
         Begin SISGalenPlus.UcDiagnosticoHIS UcDiagnosticoDetalle1 
            Height          =   3120
            Left            =   60
            TabIndex        =   48
            Top             =   375
            Width           =   11685
            _ExtentX        =   20611
            _ExtentY        =   5503
         End
         Begin VB.CommandButton btnQuitarCpt 
            DisabledPicture =   "AdmisionCEatenciones.frx":8C5F
            DownPicture     =   "AdmisionCEatenciones.frx":8FEA
            Height          =   555
            Left            =   11280
            Picture         =   "AdmisionCEatenciones.frx":937D
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Elimina CPT"
            Top             =   5280
            Width           =   615
         End
         Begin SISGalenPlus.UcPrograma UcProgramaMaterno 
            Height          =   7125
            Left            =   -74940
            TabIndex        =   40
            Top             =   360
            Width           =   11655
            _ExtentX        =   20558
            _ExtentY        =   12568
         End
         Begin VB.Frame Frame2 
            Caption         =   "Condición del paciente"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   60
            TabIndex        =   28
            Top             =   4680
            Width           =   11070
            Begin VB.ComboBox cmbIdCondicionEnElServicio 
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
               Left            =   6720
               TabIndex        =   30
               Top             =   255
               Width           =   3945
            End
            Begin VB.ComboBox cmbIdCondicionEnElEstablecimiento 
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
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   29
               Top             =   240
               Width           =   3945
            End
            Begin VB.Label Label55 
               Caption         =   "En el servicio"
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
               Left            =   5400
               TabIndex        =   32
               Top             =   240
               Width           =   1155
            End
            Begin VB.Label Label56 
               Caption         =   "En el estab."
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
               Left            =   165
               TabIndex        =   31
               Top             =   300
               Width           =   2265
            End
         End
         Begin VB.CommandButton btnCpt 
            Caption         =   "CPT"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   11280
            Picture         =   "AdmisionCEatenciones.frx":970E
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Agregar Procedimientos realizados en el Consultorio"
            Top             =   4680
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtNroHijos 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   26
            Top             =   4005
            Width           =   435
         End
         Begin VB.Frame Frame9 
            Caption         =   "Información complementaria del diagnóstico"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1110
            Left            =   60
            TabIndex        =   24
            Top             =   3555
            Width           =   8625
            Begin VB.TextBox txtCitaDxMedico 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   810
               Left            =   120
               MaxLength       =   1000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   25
               Top             =   240
               Width           =   8460
            End
         End
         Begin UltraGrid.SSUltraGrid grdOtrosCpt 
            Height          =   975
            Left            =   0
            TabIndex        =   34
            Top             =   5280
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   1720
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
            Caption         =   "Otros CPT"
         End
         Begin VB.Label lblHemoglobina 
            AutoSize        =   -1  'True
            Caption         =   "Hemoglobina"
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
            Left            =   9180
            TabIndex        =   82
            Top             =   4410
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "N° Hijos"
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
            Left            =   10545
            TabIndex        =   35
            Top             =   4050
            Width           =   645
         End
      End
      Begin VB.TextBox txtCitaExClinicos 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4845
         Left            =   435
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   41
         Top             =   960
         Width           =   8820
      End
      Begin SISGalenPlus.UcEpisodioClinico UcEpisodioClinico1 
         Height          =   1095
         Left            =   -74880
         TabIndex        =   72
         Top             =   4170
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   1931
      End
      Begin SISGalenPlus.ucCitasLista ucCitasLista1 
         Height          =   4875
         Left            =   -69510
         TabIndex        =   80
         Top             =   420
         Width           =   6285
         _ExtentX        =   11086
         _ExtentY        =   8599
      End
      Begin VB.Label lblNroAtencion 
         Alignment       =   1  'Right Justify
         Caption         =   ".."
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
         Left            =   -65280
         TabIndex        =   73
         Top             =   3240
         Width           =   1995
      End
   End
   Begin VB.Image pi_ImagSeleccionada 
      BorderStyle     =   1  'Fixed Single
      Height          =   1350
      Left            =   12360
      MouseIcon       =   "AdmisionCEatenciones.frx":9C98
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   3555
      Width           =   1560
   End
End
Attribute VB_Name = "AdmisionCEatenciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Atención del Médico
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------

Dim mRs_Productos As New ADODB.Recordset

Option Explicit
Dim lbHuboCambioEnDato As Boolean
Dim mo_Teclado As New sighEntidades.Teclado
Dim mo_Formulario As New sighEntidades.Formulario
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim ml_idUsuario As Long
Dim mb_ExistenDatos As Boolean

Public ms_Atencion As String

Dim mo_sighProxies As New SIGHProxies.Procesos
Dim mo_ReglasSISgalenhos As New SIGHSis.ReglasSISgalenhos
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminServiciosGeograficos As New SIGHNegocios.ReglasServGeograf
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion 'WCG_2006
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim ml_TipoServicio As sghTipoServicio
Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
Dim mo_Reniec As New ReniecGalenhos
'
Dim mo_Apariencia As New sighEntidades.GridInfragistic
Dim ml_TipoVistaForm As sghTipoVistaFormAtenciones
Dim ml_EstadoCuenta As Long
'
Dim mo_cmbIdDestinoAtencion As New sighEntidades.ListaDespleglable
Dim mo_cmbIdTipoReferenciaDestino As New sighEntidades.ListaDespleglable
'
Dim mo_cmbIdCondicionEnElServicio  As New sighEntidades.ListaDespleglable
Dim mo_cmbIdCondicionEnElEstablecimiento  As New sighEntidades.ListaDespleglable
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
Dim lnIdPlanSIS As Long, lcDniSIS As String, lnAfiliacionSIS1 As String, lnAfiliacionSIS2 As String
Dim lnAfiliacionSIS3 As String, lnAfiliacionSIS4 As Long, lcSIScodigo As String
Dim lcCodigoEstablecimientoAdscripcionSIS As String, lbEncontroAfiliadoEnWebSIS As Boolean
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim oRsFormaPago As New ADODB.Recordset
Dim oRsFuentesFinanciamiento As New ADODB.Recordset
Dim oRsServiciosIntermedios As New Recordset
Dim rsServicio As New Recordset
Dim oRsSoloLabActividades As New Recordset
Dim lnFormaPagoAnterior As Long
Dim lnIdFactServicios As Long

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
'
'------------------------------------------------------------------------------------
'                               VARIABLE PARA RECETAS
'------------------------------------------------------------------------------------
Dim lnRecetaRayosX As Long, lnRecetaEcografiaO As Long, lnRecetaEcografiaG As Long
Dim lnRecetaTomografia As Long, lnRecetaAnatomiaP As Long, lnRecetaPatologiaC As Long
Dim lnRecetaBancoS As Long, lnRecetaFarmacia As Long, lnRecetaOtrosCpt As Long

'                               VARIABLE PARA MODULO PROGRAMA MATERNO
'------------------------------------------------------------------------------------
Const lnProgramaMaterno As Long = 1 'IdPrograma Materno (Programas)
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
'                               VARIABLE PARA MODULO PROGRAMA PERINATAL
'------------------------------------------------------------------------------------
'mgaray20141024
Dim ml_idOrdenServicioInmunizaciones As Long

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
Dim lbElMedicoNOregistraDatosCE As String
Dim lbCargaUnaSolaVez As Boolean
Dim mo_lbEsCitaAdicional As Boolean
Const lbCargaAlaVezCitaPacienteAtencionDA As Boolean = False
Const lcPagoCita As String = "Pagada"
Dim ml_ldFechaIngreso As Date, ml_lnEdadEnDias As Long, ml_lcServicio As String, ml_IdServicio As Long
Dim ml_lcMedico As String, ml_lcHoraIngreso As String, ml_lnIdTipoEdad As Long
Dim ml_IdFuenteFinanciamiento As Long, ml_IdFormaPago As Long
Dim mb_ControlNuevoMaterno As Boolean
'***********debb-27/05/2015 (inicio)**********
Dim ml_idOrden As Long
Dim ml_idOrden_idCuenta As Long
Dim ml_AScorrelativo As Long
Dim oRsGrdOtrosCpt As New Recordset
Dim oRsTipoDx As New Recordset
Dim oRsServiciosAtenSimultaneaFuaXcorrelativo As New Recordset
Dim lcAD040 As String
Dim lnPesoKg As Double
Dim lb_YaSeRegistroDatos As Boolean
Dim ml_ups As String
Dim mc_FuaVersionFormato As String
Dim lbYaHuboDespacho As Boolean
Dim lc_HoraQueCargaFormulario As String
Dim lbTienePermisoParaRegistrarAtencionesPasadas As Boolean
Dim lbTieneLicenciaParaMensajeAcelulares As Boolean
Dim lbImpresionDeAtencionDistintoAlGrabar As Boolean
Dim lbTienePermisoParaImprimirAtencion As Boolean

Dim IdCS As Long
Dim lidatencionC As String

Property Let cidatencion(lValue As String)
   lidatencionC = lValue
End Property


Property Let cidCS(lValue As Long)
   IdCS = lValue
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
        mo_cmbIdCondicionEnElServicio.BoundColumn = "IdTipoCondicionPaciente"
        mo_cmbIdCondicionEnElServicio.ListField = "DescripcionLarga"
        Set mo_cmbIdCondicionEnElServicio.RowSource = mo_AdminServiciosComunes.TiposCondicionPacienteSeleccionarTodos
        
        mo_cmbIdCondicionEnElEstablecimiento.BoundColumn = "IdTipoCondicionPaciente"
        mo_cmbIdCondicionEnElEstablecimiento.ListField = "DescripcionLarga"
        Set mo_cmbIdCondicionEnElEstablecimiento.RowSource = mo_AdminServiciosComunes.TiposCondicionPacienteSeleccionarTodos
        
        mo_cmbIdDestinoAtencion.BoundColumn = "IdDestinoAtencion"
        mo_cmbIdDestinoAtencion.ListField = "DescripcionLarga"
        Set mo_cmbIdDestinoAtencion.RowSource = mo_AdminAdmision.TiposDestinoAtencionSeleccionarDestinosDeConsultoriosExternos

        mo_cmbIdTipoReferenciaDestino.BoundColumn = "IdTipoReferencia"
        mo_cmbIdTipoReferenciaDestino.ListField = "DescripcionLarga"
        Set mo_cmbIdTipoReferenciaDestino.RowSource = mo_AdminServiciosComunes.TiposReferenciaSeleccionarTodos

        Me.UcDiagnosticoDetalle1.TipoDiagnostico = sghAtencionConsultaExterna
        Me.UcDiagnosticoDetalle1.ConfigurarComboBoxes
        Me.UcDiagnosticoDetalle1.EditaLabConfHIS
        '
        Set cmbServicioReferenciaD.ListSource = mo_AdminServiciosComunes.SuSalud_upsSeleccionarTodos   'debb-21/06/2016
        '

End Sub

Private Sub btnAgregaApoyoDx_Click()
    Dim oReceta As New RecetaDetalle
    oReceta.Opcion = sghAgregar
    oReceta.CargaDxParaFarmacia UcDiagnosticoDetalle1.DevuelveDx
    oReceta.idTipoServicio = sghTipoServicio.sghConsultaExterna
    oReceta.idUsuario = sighEntidades.Usuario
    oReceta.idCuentaAtencion = mo_Atenciones.idCuentaAtencion
    oReceta.IdMedicoServicioActual = mo_Atenciones.IdMedicoIngreso
    oReceta.Show 1
    Set oReceta = Nothing

End Sub

Private Sub btnAntecedentesPersonales_Click()
    Dim oFormulario As New FrmAtenInteAntecedPaciente
    oFormulario.idUsuario = ml_IdPaciente
     
    oFormulario.ucHCAntecedentes1.Inicializar (ml_IdPaciente)
    oFormulario.ucHCAntecedentes1.idUsuario = ml_idUsuario
    oFormulario.Caption = Me.Caption
    
    oFormulario.Show vbModal
    
    If oFormulario.GeneroPlanIntegral = True Then
        'debb-09/06/2016 (inicio)
        If wxParametro502 = "S" Then
        Else
            Call ucPerinatal1.cargarDatosAtencionIntegral
        End If
        'Call ucPerinatal1.cargarDatosAtencionIntegral
        'debb-09/06/2016 (fin)
    End If
End Sub

Private Sub btnBuscaHistoricos_Click()
    Dim oBuscaHistoricos As New AdmisionCEhistorico
    If Me.ucPerinatalAS1.Visible = True Then
       oBuscaHistoricos.MuestraTab = 1
    End If
    oBuscaHistoricos.Paciente = lcHistoriaYpaciente
    oBuscaHistoricos.idPaciente = ml_IdPaciente
    oBuscaHistoricos.idTipoSexo = mo_paciente.idTipoSexo
    oBuscaHistoricos.NroHistoriaClinica = mo_paciente.NroHistoriaClinica  ' Val(Mid(lcHistoriaYpaciente, 2, InStr(lcHistoriaYpaciente, ")") - 2))
    oBuscaHistoricos.Show 1
    Set oBuscaHistoricos = Nothing
End Sub




'*****debb-27/05/2015
Private Sub btnCpt_Click()
    Dim oCpt As New FacOrdenServicioDetalle
    If ml_AScorrelativo = 0 Then
        'Dim orsTemp As New Recordset
        oCpt.FormMostradoDesde = 1
        oCpt.lbNOValidaCodigoPrestacion = True
        oCpt.PuntoCarga = 1   'consumo en el servicio
        'Set orsTemp = grdOtrosCpt.DataSource
        oCpt.Opcion = sghAgregar
        oCpt.idUsuario = ml_idUsuario
        oCpt.idCuentaAtencion = ml_idCuentaAtencion
        oCpt.Show 1
        CargaCPTrealizadosEnVariosServicios False
    Else
        '*******Atencion en más de 1 consultorio. Ejm CRED e INMUNIZACIONES
        Dim oAdmisionCEatencSimultanea As New AdmisionCEatencSimultanea
        Dim lnIdCta As Long
        oAdmisionCEatencSimultanea.FormLlamante = "CPT"
        oAdmisionCEatencSimultanea.Correlativo = ml_AScorrelativo
        oAdmisionCEatencSimultanea.Show 1
        lnIdCta = oAdmisionCEatencSimultanea.idCuentaAtencion
        If lnIdCta > 0 Then
           ' Dim orsTemp As New Recordset
            oCpt.FormMostradoDesde = 1
            oCpt.lbNOValidaCodigoPrestacion = True
            oCpt.PuntoCarga = 1   'consumo en el servicio
            'Set orsTemp = grdOtrosCpt.DataSource
            oCpt.Opcion = sghAgregar
            oCpt.idUsuario = ml_idUsuario
            oCpt.idCuentaAtencion = lnIdCta
            oCpt.Show 1
            CargaCPTrealizadosEnVariosServicios False
        End If
        Set oAdmisionCEatencSimultanea = Nothing
    End If
    Set oCpt = Nothing
End Sub

Sub GeneraTmpCPT()
       If oRsGrdOtrosCpt.State = 1 Then Set oRsGrdOtrosCpt = Nothing
       With oRsGrdOtrosCpt
              .Fields.Append "IdPuntoCarga", adInteger, 4, adFldIsNullable
              .Fields.Append "IdProducto", adInteger, 4, adFldIsNullable
              .Fields.Append "Codigo", adVarChar, 10, adFldIsNullable
              .Fields.Append "Nombre", adVarChar, 255, adFldIsNullable
             ' .Fields.Append "labConfHIS", adVarChar, 3, adFldIsNullable
              .Fields.Append "cantidad", adInteger, 4, adFldIsNullable
              .Fields.Append "Precio", adDouble
              .Fields.Append "Total", adDouble
              .Fields.Append "IdCuentaAtencion", adInteger, 4, adFldIsNullable
              .Fields.Append "IdOrden", adInteger, 4, adFldIsNullable
              .Fields.Append "labConfHIS", adVarChar, 3, adFldIsNullable
              .Fields.Append "Fua", adInteger
              .Fields.Append "Consultorio", adVarChar, 100, adFldIsNullable
              .Fields.Append "grupo", adInteger
              .Fields.Append "subgrupo", adInteger
              .Fields.Append "IdServicio", adInteger
              .Fields.Append "IdTipoDiagnostico", adInteger, 4, adFldIsNullable
              .Fields.Append "ups", adVarChar, 6, adFldIsNullable
              .Fields.Append "OrdenPago", adInteger
              .CursorType = adOpenKeyset
              .LockType = adLockOptimistic
              .Open
        End With
        
        

        
        
        Set grdOtrosCpt.DataSource = oRsGrdOtrosCpt
        If wxParametro302 = "S" And ml_IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
            grdOtrosCpt.Bands(0).Columns("fua").Header.Appearance.ForeColor = vbWhite
            grdOtrosCpt.Bands(0).Columns("fua").Header.Appearance.BackColor = vbRed
            grdOtrosCpt.Bands(0).Columns("fua").Header.Appearance.Font.Bold = True
            grdOtrosCpt.Bands(0).Columns("Fua").Width = 700
            grdOtrosCpt.Bands(0).Columns("Fua").Hidden = False
            Dim lnFor As Integer
            On Error Resume Next
            With grdOtrosCpt.ValueLists.Add("FuaList").ValueListItems
                 For lnFor = 1 To 20
                     .Add lnFor, "N° " & Trim(Str(lnFor))
                 Next
            End With
            grdOtrosCpt.Bands(0).Columns("Fua").ValueList = "FuaList"
            
        Else
            grdOtrosCpt.Bands(0).Columns("Fua").Hidden = True
        End If

End Sub
'*****debb-27/05/2015
Sub CargaCPTrealizadosEnVariosServicios(lbDesdeCargaDatosAlosControles As Boolean)
       Dim oRsTmp1 As New Recordset
       Dim oRsTmp2 As New Recordset
       Dim lnFua As Long, lnIdOrdenPago99 As Long
       Dim oConexion99 As New Connection
       oConexion99.CommandTimeout = 900
       oConexion99.CursorLocation = adUseClient
       oConexion99.Open sighEntidades.CadenaConexion
       If oRsGrdOtrosCpt.RecordCount > 0 Then
          Set oRsTmp1 = oRsGrdOtrosCpt.Clone()
       End If
       If ml_AScorrelativo = 0 Then
            Set oRsServiciosIntermedios = mo_AdminAdmision.BuscaAtencionesCptCEparaFormatoHIS(ml_idCuentaAtencion, sghPuntosCargaBasicos.sghPtoCargaServicioHospitalizacion)
            If oRsServiciosIntermedios.RecordCount = 0 And sighEntidades.Parametro561 = "S" Then
               oRsServiciosIntermedios.Close
               Set oRsServiciosIntermedios = mo_AdminAdmision.BuscaAtencionesCptCEparaFormatoHIS(ml_idCuentaAtencion, sghPuntosCargaBasicos.sghPtoCargaCaja)
            End If
       Else
            Set oRsServiciosIntermedios = mo_AdminAdmision.ServiciosAtenSimultaneaMovCpt(ml_AScorrelativo)
            If oRsServiciosIntermedios.RecordCount = 0 And sighEntidades.Parametro561 = "S" Then
               oRsServiciosIntermedios.Close
               Set oRsServiciosIntermedios = mo_AdminAdmision.BuscaAtencionesCptCEparaFormatoHIS(ml_AScorrelativo, sghPuntosCargaBasicos.sghPtoCargaCaja)
            End If
       End If
       '
       GeneraTmpCPT
'       If oRsGrdOtrosCpt.RecordCount > 0 Then
'          oRsGrdOtrosCpt.MoveFirst
'          Do While Not oRsGrdOtrosCpt.EOF
'             If oRsGrdOtrosCpt!Grupo = 0 Then
'                oRsGrdOtrosCpt.Delete
'             End If
'             oRsGrdOtrosCpt.MoveNext
'          Loop
'       End If
       '
       If oRsServiciosIntermedios.RecordCount > 0 Then
          oRsServiciosIntermedios.MoveFirst
          Do While Not oRsServiciosIntermedios.EOF
             '
             lnFua = 1
             If oRsTmp1.State = 1 Then
                If oRsTmp1.RecordCount > 0 Then
                   oRsTmp1.MoveFirst
                   Do While Not oRsTmp1.EOF
                      If oRsTmp1!idProducto = oRsServiciosIntermedios!idProducto And oRsTmp1!IdOrden = oRsServiciosIntermedios!IdOrden Then
                         lnFua = oRsTmp1!FUA
                         Exit Do
                      End If
                      oRsTmp1.MoveNext
                   Loop
                End If
             End If
             '
             lnIdOrdenPago99 = 0
             If mo_Atenciones.IdFormaPago = 1 Then
                Set oRsTmp2 = mo_AdminFacturacion.FactOrdenServicioPagosSeleccionarPorIdOrden(oRsServiciosIntermedios!IdOrden, oConexion99)
                If oRsTmp2.RecordCount > 0 Then
                   lnIdOrdenPago99 = oRsTmp2!IdOrdenPago
                End If
                oRsTmp2.Close
             End If
             '
             oRsGrdOtrosCpt.AddNew
             oRsGrdOtrosCpt.Fields!idPuntoCarga = oRsServiciosIntermedios!idPuntoCarga
             oRsGrdOtrosCpt.Fields!idProducto = oRsServiciosIntermedios!idProducto
             If ml_AScorrelativo = 0 Then
                oRsGrdOtrosCpt.Fields!Consultorio = ml_lcServicio
             Else
                oRsGrdOtrosCpt.Fields!Consultorio = oRsServiciosIntermedios!Consultorio
             End If
             oRsGrdOtrosCpt.Fields!Codigo = oRsServiciosIntermedios!Codigo
             oRsGrdOtrosCpt.Fields!nombre = oRsServiciosIntermedios!nombre
             oRsGrdOtrosCpt.Fields!labConfHIS = oRsServiciosIntermedios!labConfHIS
             oRsGrdOtrosCpt.Fields!Cantidad = oRsServiciosIntermedios!Cantidad
             oRsGrdOtrosCpt.Fields!precio = oRsServiciosIntermedios!precio
             oRsGrdOtrosCpt.Fields!Total = oRsServiciosIntermedios!Total
             oRsGrdOtrosCpt.Fields!idCuentaAtencion = oRsServiciosIntermedios!idCuentaAtencion
             oRsGrdOtrosCpt.Fields!IdOrden = oRsServiciosIntermedios!IdOrden
             oRsGrdOtrosCpt.Fields!FUA = lnFua
             oRsGrdOtrosCpt.Fields!IdServicio = oRsServiciosIntermedios!IdServicio
             oRsGrdOtrosCpt.Fields!Grupo = IIf(IsNull(oRsServiciosIntermedios!grupoHIS), 0, oRsServiciosIntermedios!grupoHIS)
             oRsGrdOtrosCpt.Fields!SubGrupo = IIf(IsNull(oRsServiciosIntermedios!subgrupoHIS), 0, oRsServiciosIntermedios!subgrupoHIS)
             oRsGrdOtrosCpt.Fields!UPS = mo_AdminAdmision.BuscaUPSactualDelPaciente(oRsServiciosIntermedios!IdServicio)
             oRsGrdOtrosCpt.Fields!OrdenPago = lnIdOrdenPago99
             oRsGrdOtrosCpt.Update
             oRsServiciosIntermedios.MoveNext
          Loop
          If lbDesdeCargaDatosAlosControles = True And oRsGrdOtrosCpt.RecordCount > 0 Then
                oRsServiciosAtenSimultaneaFuaXcorrelativo.Filter = "idtipo=1"
                If oRsServiciosAtenSimultaneaFuaXcorrelativo.RecordCount > 0 Then
                   oRsServiciosAtenSimultaneaFuaXcorrelativo.MoveFirst
                   Do While Not oRsServiciosAtenSimultaneaFuaXcorrelativo.EOF
                      oRsGrdOtrosCpt.MoveFirst
                      Do While Not oRsGrdOtrosCpt.EOF
                         If oRsGrdOtrosCpt.Fields!idProducto = oRsServiciosAtenSimultaneaFuaXcorrelativo!Item Then
                             oRsGrdOtrosCpt.Fields!FUA = oRsServiciosAtenSimultaneaFuaXcorrelativo!idFuaCorrelativo
                         End If
                         oRsGrdOtrosCpt.MoveNext
                      Loop
                      oRsServiciosAtenSimultaneaFuaXcorrelativo.MoveNext
                   Loop
                End If
          End If
       End If
       If oRsGrdOtrosCpt.RecordCount > 0 Then
          oRsGrdOtrosCpt.MoveFirst
       End If
       Set grdOtrosCpt.DataSource = oRsGrdOtrosCpt
       Set oRsTmp1 = Nothing
       If wxParametro302 = "S" And mo_Atenciones.IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
          grdOtrosCpt.Bands(0).Columns("fua").Hidden = False
       Else
          grdOtrosCpt.Bands(0).Columns("fua").Hidden = True
       End If
       If ml_AScorrelativo = 0 Then
          grdOtrosCpt.Bands(0).Columns("consultorio").Hidden = True
       Else
          grdOtrosCpt.Bands(0).Columns("consultorio").Hidden = False
       End If
       oConexion99.Close
       Set oRsTmp2 = Nothing
       Set oConexion99 = Nothing
End Sub

Private Sub btnguardarcs_Click()
If ValidarDatosObligatoriosCS() Then
If AgregarDatosCS() Then
                MsgBox "Se registrarón correctamente los datos " + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
                grdServicios.LimpiarGrilla
            Else
                MsgBox "No se registrarón los datos " + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
            End If
       End If
End Sub


'Sub CargaCPTrealizadosEnElServicio()
'    Set oRsServiciosIntermedios = mo_AdminAdmision.BuscaAtencionesCptCEparaFormatoHIS(ml_idCuentaAtencion)
'    Set grdOtrosCpt.DataSource = oRsServiciosIntermedios
'End Sub

Private Sub btnImprimeAtencion_Click()
    On Error GoTo ErrImp
    
    Dim oRptHistoriaClinicaCE As New RptHistoriaClinicaCE
    Dim oDODiagnostico As New DODiagnostico, lcDx As String
    Dim lcFechaConsulta As Date
    Dim lcDxMedico As String
    Dim orsOtrosProcedDX As New Recordset
    Dim lcOtrosCptDxMedicos As String
    Set oDODiagnostico = mo_AdminFacturacion.DevuelveDxAltaMedica(mo_Atenciones.idAtencion, 1)
    lcDx = Left(Trim(oDODiagnostico.CodigoCIE2004) & " " & Trim(oDODiagnostico.descripcion), 100)
    lcDx = mo_AdminFacturacion.DevuelveDxAltaMedicaTodosDx(mo_Atenciones.idAtencion, 1, "")
    
    lcFechaConsulta = ml_ldFechaIngreso
    lcDxMedico = Mid(txtCitaDxMedico.Text, InStr(txtCitaDxMedico.Text, lcLineaChar) + 1, 1000)
    txtCitaExClinicos.Text = Me.UcRecetas1.DevuelveRecetaAntesDeImprimir
    
    Set orsOtrosProcedDX = grdOtrosCpt.DataSource 'Actualizado 31032015
    lcOtrosCptDxMedicos = ""
    If orsOtrosProcedDX.RecordCount > 0 Then
        orsOtrosProcedDX.MoveFirst
        Do While Not orsOtrosProcedDX.EOF
             lcOtrosCptDxMedicos = lcOtrosCptDxMedicos & "[" & orsOtrosProcedDX!Codigo & " = " & orsOtrosProcedDX!nombre & "], "
             orsOtrosProcedDX.MoveNext
        Loop
    End If
    
    'mgaray201410d
    Dim sDescripcionExamenes As String, sDescripcionRecetas As String
    sDescripcionExamenes = Me.UcRecetas1.DevuelveSoloExamenesParaImpresion
    sDescripcionRecetas = Me.UcRecetas1.DevuelveSoloRecetaParaImpresion
    
    Dim oDOAtencionesCE As DOAtencionesCE
    Set oDOAtencionesCE = RetornaObjetoDatosTriaje()
    'debb2014d
    If lbElConsultorioUsaModuloMaterno = True And mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE Then
        oRptHistoriaClinicaCE.CrearReporteCeAtencionPacienteConPrograma ml_idAtencion, _
                              lcHistoriaYpaciente & " (Edad: " & Trim(Str(ml_lnEdadEnDias)) & ")", _
                              ml_idCuentaAtencion, ml_lcServicio, lcFechaConsulta, ml_lcMedico, _
                              oDOAtencionesCE.TriajePresion, oDOAtencionesCE.triajeTalla, oDOAtencionesCE.TriajeTemperatura, oDOAtencionesCE.triajePeso, _
                              txtCitaMotivo.Text, txtCitaExamenClinico.Text, lcDxMedico, lcDx, _
                              TxtCitaTratamiento.Text, sDescripcionExamenes, txtCitaObservaciones.Text, True, _
                              lnProgramaMaterno, UcProgramaMaterno.EsControlActual, _
                              UcProgramaMaterno.DevuelveNumeroControl, UcProgramaMaterno.DevuelveFechaControl, _
                              "Programa Materno", UcProgramaMaterno.DevuelveDatosProCabecera, _
                              UcProgramaMaterno.DevuelveDatosProControles, UcProgramaMaterno.DevuelveProDiagnosticos, _
                              UcProgramaMaterno.DevuelveProProcedimientos, UcProgramaMaterno.DevuelveProTratamientos, Me.hwnd, sDescripcionRecetas, lcOtrosCptDxMedicos
    ElseIf lbElConsultorioUsaModuloPerinatal = True And mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE Then
        'debb-09/06/2016 (inicio)
        If wxParametro502 = "S" Then
            oRptHistoriaClinicaCE.CrearReporteCeAtencionPaciente Me.hwnd, ml_idAtencion, _
                              lcHistoriaYpaciente & " (Edad: " & Trim(Str(ml_lnEdadEnDias)) & ")", _
                              ml_idCuentaAtencion, ml_lcServicio, lcFechaConsulta, ml_lcMedico, _
                              oDOAtencionesCE.TriajePresion, oDOAtencionesCE.triajeTalla, oDOAtencionesCE.TriajeTemperatura, oDOAtencionesCE.triajePeso, _
                              txtCitaMotivo.Text, txtCitaExamenClinico.Text, lcDxMedico, lcDx, _
                              TxtCitaTratamiento.Text, txtCitaExClinicos.Text, txtCitaObservaciones.Text, _
                               False, ml_IdPaciente, _
                              lbImpresionDeAtencionDistintoAlGrabar, _
                              IIf(Me.ucPerinatalAS1.Visible = True, Me.ucPerinatalAS1.DevuelveCptInmunizaciones, Nothing), _
                              IIf(Me.ucPerinatalAS1.Visible = True, Me.ucPerinatalAS1.DevuelveCptFrecuentes, Nothing), _
                              IIf(Me.ucPerinatalAS1.Visible = True, Me.ucPerinatalAS1.DevuelveDxDesarrollo, Nothing), _
                              IIf(Me.ucPerinatalAS1.Visible = True, Me.ucPerinatalAS1.DevuelveDxMorbilidad, Nothing), _
                              IIf(Me.ucPerinatalAS1.Visible = True, Me.ucPerinatalAS1.DevuelveMedicamentos, Nothing), _
                              sDescripcionRecetas, lcOtrosCptDxMedicos
        Else
             oRptHistoriaClinicaCE.CrearReporteCeAtencionPacienteCRED Me.hwnd, ml_idAtencion, _
                              lcHistoriaYpaciente & " (Edad: " & Trim(Str(ml_lnEdadEnDias)) & ")", _
                              ml_idCuentaAtencion, ml_lcServicio, lcFechaConsulta, ml_lcMedico, _
                              oDOAtencionesCE.TriajePresion, oDOAtencionesCE.triajeTalla, oDOAtencionesCE.TriajeTemperatura, oDOAtencionesCE.triajePeso, _
                              txtCitaMotivo.Text, txtCitaExamenClinico.Text, lcDxMedico, lcDx, _
                              TxtCitaTratamiento.Text, sDescripcionExamenes, txtCitaObservaciones.Text, _
                              IIf(Me.ucPerinatal1.Visible = True, Me.ucPerinatal1.DevuelveCptInmunizaciones, Nothing), _
                              IIf(Me.ucPerinatal1.Visible = True, Me.ucPerinatal1.DevuelveCptFrecuentes, Nothing), _
                              IIf(Me.ucPerinatal1.Visible = True, Me.ucPerinatal1.DevuelveDxDesarrollo, Nothing), _
                              IIf(Me.ucPerinatal1.Visible = True, Me.ucPerinatal1.DevuelveDxMorbilidad, Nothing), _
                              IIf(Me.ucPerinatal1.Visible = True, Me.ucPerinatal1.DevuelveMedicamentos, Nothing), _
                              sDescripcionRecetas, lcOtrosCptDxMedicos
        End If
'             oRptHistoriaClinicaCE.CrearReporteCeAtencionPacienteCRED Me.hwnd, ml_idAtencion, _
'                              lcHistoriaYpaciente & " (Edad: " & Trim(Str(ml_lnEdadEnDias)) & ")", _
'                              ml_idCuentaAtencion, ml_lcServicio, lcFechaConsulta, ml_lcMedico, _
'                              oDOAtencionesCE.TriajePresion, oDOAtencionesCE.TriajeTalla, oDOAtencionesCE.TriajeTemperatura, oDOAtencionesCE.TriajePeso, _
'                              txtCitaMotivo.Text, txtCitaExamenClinico.Text, lcDxMedico, lcDx, _
'                              TxtCitaTratamiento.Text, sDescripcionExamenes, txtCitaObservaciones.Text, _
'                              IIf(Me.ucPerinatal1.Visible = True, Me.ucPerinatal1.DevuelveCptInmunizaciones, Nothing), _
'                              IIf(Me.ucPerinatal1.Visible = True, Me.ucPerinatal1.DevuelveCptFrecuentes, Nothing), _
'                              IIf(Me.ucPerinatal1.Visible = True, Me.ucPerinatal1.DevuelveDxDesarrollo, Nothing), _
'                              IIf(Me.ucPerinatal1.Visible = True, Me.ucPerinatal1.DevuelveDxMorbilidad, Nothing), _
'                              IIf(Me.ucPerinatal1.Visible = True, Me.ucPerinatal1.DevuelveMedicamentos, Nothing), _
'                              sDescripcionRecetas, lcOtrosCptDxMedicos
        'debb-09/06/2016 (fin)
    Else
        oRptHistoriaClinicaCE.CrearReporteCeAtencionPaciente Me.hwnd, ml_idAtencion, _
                              lcHistoriaYpaciente & " (Edad: " & Trim(Str(ml_lnEdadEnDias)) & ")", _
                              ml_idCuentaAtencion, ml_lcServicio, lcFechaConsulta, ml_lcMedico, _
                              oDOAtencionesCE.TriajePresion, oDOAtencionesCE.triajeTalla, _
                              oDOAtencionesCE.TriajeTemperatura, oDOAtencionesCE.triajePeso, _
                              txtCitaMotivo.Text, txtCitaExamenClinico.Text, lcDxMedico, lcDx, _
                              TxtCitaTratamiento.Text, sDescripcionExamenes & sDescripcionRecetas, _
                              txtCitaObservaciones.Text, False, _
                              ml_IdPaciente, lbImpresionDeAtencionDistintoAlGrabar, _
                              , , , , , , lcOtrosCptDxMedicos
    End If
    'debb2014d
    Set oRptHistoriaClinicaCE = Nothing
    Set oDODiagnostico = Nothing
ErrImp:
    lbImpresionDeAtencionDistintoAlGrabar = True
End Sub

Private Sub btnImprimeFichaSIS_Click()
    Dim ml_FuaTipoAnexo2015 As Integer
    If mi_Opcion <> sghAgregar Then
       CargaDatosAlObjetosDeDatos
    End If
    If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE Then
         If ValidarDatosObligatorios = False Or ValidarReglas = False Then
               Exit Sub
         End If
    ElseIf mi_Opcion = sghModificar Then
         If ValidarDatosObligatorios = False Or ValidarReglas = False Or lbElMedicoNOregistraDatosCE = "N" Then
               Exit Sub
         End If
    End If
    If wxParametro553 = "S" And mo_Atenciones.IdDestinoAtencion = 11 Then
       If MsgBox("El FUA debe imprimirse al terminar su Hospitalización" & Chr(13) & _
                 "      siempre y cuando se hospitalize hoy mismo.     " & Chr(13) & _
                 "         El PACIENTE SERA HOSPITALIZADO HOY ?        ", vbQuestion + vbYesNo, "") = vbYes Then
          Exit Sub
       End If
    End If
    Dim oFua As New SIGHSis.clFUA
    oFua.idCuentaAtencion = ml_idCuentaAtencion 'mo_Atenciones.idCuentaAtencion
    oFua.lcNombrePc = mo_lcNombrePc
    oFua.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
    oFua.idUsuario = ml_idUsuario
    oFua.Opcion = mi_Opcion
    oFua.IdServicio = ml_IdServicio
    oFua.MostrarFormulario
    Set oFua = Nothing
End Sub

Private Sub Calendario_Change(ByVal NewDate As Date)
   lblProximaCita.Caption = Format(NewDate, "dd/mm/yyyy")
End Sub











Private Sub cmbIdCondicionEnElEstablecimiento_Click()
lbHuboCambioEnDato = True
End Sub



Private Sub cmbIdCondicionEnElEstablecimiento_LostFocus()
    If lbHuboCambioEnDato = True Then
      sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, cmbIdCondicionEnElEstablecimiento.Text
      lbHuboCambioEnDato = False
    End If
End Sub

Private Sub cmbIdCondicionEnElServicio_Click()
lbHuboCambioEnDato = True
End Sub

Private Sub cmbIdCondicionEnElServicio_LostFocus()
    If lbHuboCambioEnDato = True Then
      sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, cmbIdCondicionEnElServicio.Text
      lbHuboCambioEnDato = False
    End If
End Sub

Private Sub cmbServicioReferenciaD_Click()
lbHuboCambioEnDato = True
End Sub

Private Sub cmbServicioReferenciaD_LostFocus()
    If lbHuboCambioEnDato = True Then
      sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, cmbServicioReferenciaD.Text
      lbHuboCambioEnDato = False
    End If
End Sub

Private Sub cmbTipoEdad1_Click()
lbHuboCambioEnDato = True
End Sub

Private Sub cmbTipoEdad1_LostFocus()
    If lbHuboCambioEnDato = True Then
      sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, cmbTipoEdad1.Text
      lbHuboCambioEnDato = False
    End If
End Sub



Private Sub cmdGenerarPlanAtencion_Click()
    If MsgBox("Desea Generar y/o Actualizar el Plan de Atención Integral del Paciente", vbQuestion + vbYesNo, "Módulo Niño Sano") = vbNo Then
        Exit Sub
    End If
    Dim oFormulario As New FrmAtenInteAntecedPaciente
    oFormulario.idUsuario = ml_IdPaciente
     
    oFormulario.ucHCAntecedentes1.Inicializar (ml_IdPaciente)
    oFormulario.ucHCAntecedentes1.idUsuario = ml_idUsuario
    oFormulario.Caption = Me.Caption
    
    'debb-09/06/2016 (inicio)
    If wxParametro502 = "S" Then
    Else
        If oFormulario.GenerarPlanDeAtencionIntegral = True Then
            MsgBox "Plan de Atención Integral del Paciente Generado", vbInformation, "Modulo Niño Sano"
            Call ucPerinatal1.cargarDatosAtencionIntegral
        End If
    End If
'    If oFormulario.GenerarPlanDeAtencionIntegral = True Then
'        MsgBox "Plan de Atención Integral del Paciente Generado", vbInformation, "Modulo Niño Sano"
'        Call ucPerinatal1.cargarDatosAtencionIntegral
'    End If
    'debb-09/06/2016 (fin)
End Sub

Private Sub cmdImpresionReferencias_Click()
    Dim sCodigoDestino As String
    If cmbIdDestinoAtencion.Text = "" Then
       Exit Sub
    End If
    
    sCodigoDestino = Trim(Split(cmbIdDestinoAtencion.Text, " = ")(0))
    If sCodigoDestino = "R" Or sCodigoDestino = "C" Then
       If ValidarDatosObligatorios() Then
            CargaDatosAlObjetosDeDatos
            If ValidarReglas() Then
                Dim oReferencias As New SIGHReportes.clReferencias
                If sCodigoDestino = "C" Then
                   oReferencias.CreaReporteContrarefencias txtNombreDestinoReferencia.Text, mo_paciente, mo_Atenciones, _
                                                           mo_DoAtencionDatosAdicionales, False
                Else
                   oReferencias.CrearReporteReferencias txtNombreDestinoReferencia.Text, mo_paciente, mo_Atenciones, _
                                                        mo_DoAtencionDatosAdicionales, False
                End If
                Set oReferencias = Nothing
             End If
        End If
    End If
End Sub

Private Sub cmdMuestraActividades_Click()
        Dim oRsActividades As New Recordset
        Dim oRsTmp1 As New Recordset, oRsTmp2 As New Recordset, oRsDx As New Recordset
        Dim lcNombre As String, lnGrupo As Integer, lnSubGrupo As Integer, lbPrimerReg As Boolean
        Dim lnEdad As Long
        Dim lcEligio As Boolean, lcEligioLab As String, lnEligioTipo As Integer, lnEligioUPS As Long
        Dim ln_IdCuentaAtencion As Long, ln_IdOrden As Long, ln_Fua As Integer, lc_Consultorio As String
        Dim ln_idServicio As Long, lc_FuaCodigoPrestacion As String, lbUnaSolaVez As Boolean
        Dim lc_id As String, lnPrecioUnitario As Double, ln_idServicioPaciente As Long
        Dim oFactOrdenServicio As New FactOrdenServicio
        Dim oDOFactOrdenServicio As New DoFactOrdenServ
        Dim mrs_FacturacionProductos As New Recordset
        Dim oDoCatalogoServicioHosp As New DOFinanciamientoCatalogoServ
        With oRsActividades
              .Fields.Append "GrupoTIT", adVarChar, 3, adFldIsNullable
              .Fields.Append "Grupo", adInteger
              .Fields.Append "SubGrupo", adInteger
              .Fields.Append "lab", adVarChar, 3, adFldIsNullable
              .Fields.Append "Tipo", adVarChar, 20, adFldIsNullable
              .Fields.Append "id", adVarChar, 20, adFldIsNullable
              .Fields.Append "Nombre", adVarChar, 255, adFldIsNullable
              .Fields.Append "Elija", adBoolean
              .Fields.Append "ElijaTipo", adInteger
              .Fields.Append "ElijaUPS", adInteger
              .Fields.Append "ElijaLab", adVarChar, 3, adFldIsNullable
              .Fields.Append "IdCuentaAtencion", adInteger, 4, adFldIsNullable
              .Fields.Append "IdOrden", adInteger, 4, adFldIsNullable
              .Fields.Append "Fua", adInteger
              .Fields.Append "Consultorio", adVarChar, 100, adFldIsNullable
              .Fields.Append "IdServicio", adInteger
              .Fields.Append "FuaCodigoPrestacion", adVarChar, 3, adFldIsNullable
              .Fields.Append "idTipo", adInteger
              .Fields.Append "idServicioPaciente", adInteger
              .CursorType = adOpenKeyset
              .LockType = adLockOptimistic
              .Open
        End With
        Set oRsTmp1 = mo_AdminAdmision.ServiciosAtenSimultaneaImpHISxUPS(mo_AdminAdmision.BuscaUPSactualDelPaciente(mo_Atenciones.IdServicioIngreso))

        Dim oEdad As Edad, lbContiuar9 As Boolean
        '
       ' oEdad = calcularEdadDisgregada(mo_paciente.FechaNacimiento, mo_Atenciones.FechaIngreso)
        Select Case cmbTipoEdad1.ListIndex
        Case 0   'año
             oEdad.EdadAnio = Val(txtEdad1.Text)
        Case 1   'meses
             oEdad.EdadMes = Val(txtEdad1.Text)
        Case Else
             oEdad.EdadDia = Val(txtEdad1.Text)
        End Select
        oEdad.TipoEdad = cmbTipoEdad1.ListIndex + 1
        '
        If oEdad.EdadAnio > 0 Then
           oRsTmp1.Filter = "idtipoedad=1 and edadinicio" & IIf(ml_ups = "301202", "=", "<=") & oEdad.EdadAnio
        ElseIf oEdad.EdadMes > 0 Then
           oRsTmp1.Filter = "idtipoedad=2 and edadinicio=" & oEdad.EdadMes
        Else
          oRsTmp1.Filter = "idtipoedad=3"
        End If
        If oRsTmp1.RecordCount > 0 Then
           Set oRsDx = Me.UcDiagnosticoDetalle1.DevuelveDx
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
              lnGrupo = oRsTmp1!Grupo

              lbPrimerReg = True
              Do While Not oRsTmp1.EOF And lnGrupo = oRsTmp1!Grupo
                 lbContiuar9 = True
                 If oEdad.EdadDia > 0 And oEdad.EdadAnio = 0 And oEdad.EdadMes = 0 Then
                    If Not (oEdad.EdadDia >= oRsTmp1!EdadInicio And oEdad.EdadDia <= oRsTmp1!EdadFinal) Then
                       lbContiuar9 = False
                    End If
                 End If

                 If (lnPesoKg >= oRsTmp1!PesoKgMenor And lnPesoKg <= oRsTmp1!PesoKgMayor) And lbContiuar9 = True Then
                    lnSubGrupo = oRsTmp1!subgrupoOrden
                    lcNombre = ""
                    Select Case oRsTmp1!idTipo
                    Case sghActividadesTipo.TipoCPT
                            Set oRsTmp2 = mo_AdminCaja.FactCatalogoServiciosSeleccionarPorCodigoOnombre(oRsTmp1!cpt_dx, "")
                            If oRsTmp2.RecordCount > 0 Then
                               lcNombre = Left(oRsTmp2!nombre, 255)
                            End If
                            oRsTmp2.Close
                    Case sghActividadesTipo.TipoDX
                            Set oRsTmp2 = mo_AdminServiciosComunes.DiagnosticosSeleccionarXCodigo(oRsTmp1!cpt_dx)
                            If oRsTmp2.RecordCount > 0 Then
                               lcNombre = Left(oRsTmp2!descripcion, 255)
                            End If
                             oRsTmp2.Close
                    End Select
                    '
                    lcEligio = False
                    lcEligioLab = ""
                    lnEligioTipo = 102
                    lnEligioUPS = ml_idCuentaAtencion    'mo_Atenciones.IdServicioIngreso
                    ln_IdCuentaAtencion = ml_idCuentaAtencion: ln_IdOrden = 0: ln_Fua = 0: lc_Consultorio = ml_lcServicio
                    ln_idServicio = ml_idCuentaAtencion: lc_FuaCodigoPrestacion = "": ln_idServicioPaciente = mo_Atenciones.IdServicioIngreso
                    If oRsTmp1!idTipo = 1 Then
                        oRsGrdOtrosCpt.Filter = "grupo=" & lnGrupo & " and subgrupo=" & lnSubGrupo & _
                                                " and codigo='" & Trim(oRsTmp1!cpt_dx) & "'"
                        If oRsGrdOtrosCpt.RecordCount > 0 Then
                            lcEligio = True
                            lcEligioLab = oRsGrdOtrosCpt!labConfHIS
                            lnEligioTipo = IIf(IsNull(oRsGrdOtrosCpt!idTipoDiagnostico), 102, oRsGrdOtrosCpt!idTipoDiagnostico)
                            lnEligioUPS = oRsGrdOtrosCpt!idCuentaAtencion
                            ln_IdCuentaAtencion = oRsGrdOtrosCpt!idCuentaAtencion
                            ln_IdOrden = oRsGrdOtrosCpt!IdOrden
                            ln_Fua = oRsGrdOtrosCpt!FUA
                            lc_Consultorio = oRsGrdOtrosCpt!Consultorio
                            ln_idServicio = oRsGrdOtrosCpt!IdServicio
                            ln_idServicioPaciente = oRsGrdOtrosCpt!IdServicio
                        End If
                    Else
                        oRsDx.Filter = "grupo=" & lnGrupo & " and subgrupo=" & lnSubGrupo & _
                                       " and CodigoCIE2004='" & Trim(oRsTmp1!cpt_dx) & "'"
                        If oRsDx.RecordCount > 0 Then
                            lcEligio = True
                            lcEligioLab = IIf(IsNull(oRsDx!labConfHIS), "", oRsDx!labConfHIS)
                            lnEligioTipo = oRsDx!idTipoDiagnostico
                            lnEligioUPS = oRsDx!idCuentaAtencion
                            ln_IdCuentaAtencion = oRsDx!idCuentaAtencion
                            ln_Fua = oRsDx!FUA
                            lc_Consultorio = oRsDx!Consultorio
                            ln_idServicio = oRsDx!idCuentaAtencion
                            lc_FuaCodigoPrestacion = IIf(IsNull(oRsDx!FuaCodigoPrestacion), "", oRsDx!FuaCodigoPrestacion)
                            ln_idServicioPaciente = oRsDx!IdServicio
                        End If

                    End If
                    '
                    oRsActividades.AddNew
                    If lbPrimerReg = True Then
                       lbPrimerReg = False
                       oRsActividades!GrupoTIT = Trim(Str(lnGrupo))
                    Else
                       oRsActividades!GrupoTIT = ""
                    End If
                    oRsActividades!Grupo = lnGrupo
                    oRsActividades!SubGrupo = lnSubGrupo
                    oRsActividades!lab = IIf(IsNull(oRsTmp1!lab), " ", oRsTmp1!lab)
                    oRsActividades!ID = oRsTmp1!cpt_dx
                    oRsActividades!tipo = oRsTmp1!dTipo
                    oRsActividades!nombre = lcNombre
                    oRsActividades!elija = lcEligio
                    oRsActividades!elijaTipo = lnEligioTipo - 100
                    oRsActividades!ElijaUPS = lnEligioUPS
                    oRsActividades!ElijaLab = lcEligioLab
                    oRsActividades!idCuentaAtencion = ln_IdCuentaAtencion
                    oRsActividades!IdOrden = ln_IdOrden
                    oRsActividades!FUA = ln_Fua
                    oRsActividades!Consultorio = lc_Consultorio
                    oRsActividades!IdServicio = ln_idServicio
                    oRsActividades!FuaCodigoPrestacion = lc_FuaCodigoPrestacion
                    oRsActividades!idTipo = oRsTmp1!idTipo
                    oRsActividades!idServicioPaciente = ln_idServicioPaciente
                    oRsActividades.Update
                 End If
                 oRsTmp1.MoveNext
                 If oRsTmp1.EOF Then
                    Exit Do
                 End If
              Loop
           Loop

           oRsGrdOtrosCpt.Filter = ""
           If oRsGrdOtrosCpt.RecordCount > 0 Then
              oRsGrdOtrosCpt.MoveFirst
           End If
           oRsDx.Filter = ""
           If oRsDx.RecordCount > 0 Then
              oRsDx.MoveFirst
           End If
        End If
        oRsTmp1.Close
        If oRsActividades.RecordCount > 0 Then
            Dim oAdmisionCEatencSimultanea As New AdmisionCEatencSimultanea
            Dim oRsItemsElegidos As New Recordset
            oAdmisionCEatencSimultanea.FormLlamante = "ACTIVIDADES"
            Set oAdmisionCEatencSimultanea.oRsFua = oRsActividades
            Set oAdmisionCEatencSimultanea.oRsItemsElegidos = oRsTipoDx
            oAdmisionCEatencSimultanea.Show 1
            If oAdmisionCEatencSimultanea.idCuentaAtencion = 1 Then
               Set oRsItemsElegidos = oAdmisionCEatencSimultanea.ItemsMasivosElegidos
               ActividadesHIS oRsItemsElegidos
            End If
         End If


        Set oRsActividades = Nothing
        Set oRsTmp1 = Nothing
        Set oRsDx = Nothing
        Set oFactOrdenServicio = Nothing
        Set oDOFactOrdenServicio = Nothing
        Set mrs_FacturacionProductos = Nothing
        Set oDoCatalogoServicioHosp = Nothing

End Sub

Sub CreaYllenaTemporalesActividades(lbSoloCreaTemporal As Boolean)
     If oRsSoloLabActividades.State = 1 Then oRsSoloLabActividades.Close
     With oRsSoloLabActividades
        .Fields.Append "Grupo", adInteger
        .Fields.Append "SubGrupo", adInteger
        .Fields.Append "lab", adVarChar, 3, adFldIsNullable
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
     End With
     If lbSoloCreaTemporal = False Then
        Dim oRsTmp1 As New Recordset
        Set oRsTmp1 = mo_AdminAdmision.AtencionesLabSinDxSinCptSeleccionarPorId(mo_Atenciones.idAtencion)
        If oRsTmp1.RecordCount > 0 Then
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
              oRsSoloLabActividades.AddNew
              oRsSoloLabActividades!Grupo = oRsTmp1!Grupo
              oRsSoloLabActividades!SubGrupo = oRsTmp1!SubGrupo
              oRsSoloLabActividades!lab = oRsTmp1!lab
              oRsSoloLabActividades.Update
              oRsTmp1.MoveNext
           Loop
        End If
        oRsTmp1.Close
        Set oRsTmp1 = Nothing
     End If
End Sub

Private Sub cmdRegistraActividades_Click()
        Dim lcMensajeLicencia As String
'        If  False Then  'licencia
'           Exit Sub
'        End If
        
        
        Dim oRsActividades As New Recordset
        Dim oRsTmp1 As New Recordset, oRsTmp2 As New Recordset, oRsDx As New Recordset
        Dim lcNombre As String, lnGrupo As Integer, lnSubGrupo As Integer, lbPrimerReg As Boolean
        Dim lnEdad As Long
        Dim lcEligio As Boolean, lcEligioLab As String, lnEligioTipo As Integer, lnEligioUPS As Long
        Dim ln_IdCuentaAtencion As Long, ln_IdOrden As Long, ln_Fua As Integer, lc_Consultorio As String
        Dim ln_idServicio As Long, lc_FuaCodigoPrestacion As String, lbUnaSolaVez As Boolean
        Dim lc_id As String, lnPrecioUnitario As Double, ln_idServicioPaciente As Long
        Dim oFactOrdenServicio As New FactOrdenServicio
        Dim oDOFactOrdenServicio As New DoFactOrdenServ
        Dim mrs_FacturacionProductos As New Recordset
        Dim oDoCatalogoServicioHosp As New DOFinanciamientoCatalogoServ
        With oRsActividades
              .Fields.Append "GrupoTIT", adVarChar, 3, adFldIsNullable
              .Fields.Append "Grupo", adInteger
              .Fields.Append "SubGrupo", adInteger
              .Fields.Append "lab", adVarChar, 3, adFldIsNullable
              .Fields.Append "Tipo", adVarChar, 20, adFldIsNullable
              .Fields.Append "idTipo", adInteger
              .Fields.Append "id", adVarChar, 20, adFldIsNullable
              .Fields.Append "Nombre", adVarChar, 255, adFldIsNullable
              .Fields.Append "Elija", adBoolean
              .Fields.Append "ElijaTipo", adInteger
              .Fields.Append "ElijaUPS", adInteger
              .Fields.Append "ElijaLab", adVarChar, 3, adFldIsNullable
              .Fields.Append "IdCuentaAtencion", adInteger, 4, adFldIsNullable
              .Fields.Append "IdOrden", adInteger, 4, adFldIsNullable
              .Fields.Append "Fua", adInteger
              .Fields.Append "Consultorio", adVarChar, 100, adFldIsNullable
              .Fields.Append "IdServicio", adInteger
              .Fields.Append "FuaCodigoPrestacion", adVarChar, 3, adFldIsNullable
              .Fields.Append "idServicioPaciente", adInteger
              .CursorType = adOpenKeyset
              .LockType = adLockOptimistic
              .Open
        End With
        
        'Set oRsTmp1 = mo_AdminAdmision.ServiciosAtenSimultaneaImpHISxUPS(mo_AdminAdmision.BuscaUPSactualDelPaciente(mo_Atenciones.IdServicioIngreso))
        With oRsTmp1
              .Fields.Append "Grupo", adInteger
              .Fields.Append "subGrupoOrden", adInteger
              .Fields.Append "EdadInicio", adInteger
              .Fields.Append "EdadFinal", adInteger
              .Fields.Append "PesoKgMenor", adDouble
              .Fields.Append "PesoKgMayor", adDouble
              .Fields.Append "idTipo", adInteger
              .Fields.Append "dTipo", adVarChar, 30
              .Fields.Append "cpt_dx", adVarChar, 20
              .Fields.Append "lab", adVarChar, 3
              .CursorType = adOpenKeyset
              .LockType = adLockOptimistic
              .Open
        End With
        
        
        
        Dim lnFor1 As Integer, lnFor2 As Integer, lnIdTipo3 As Integer, lcCpt_Dx3 As String, lcDtipo3 As String, lcLab3 As String
        Set oRsDx = Me.UcDiagnosticoDetalle1.DevuelveDx
        For lnFor1 = 1 To 5
            For lnFor2 = 1 To 6
                lnIdTipo3 = sghActividadesTipo.TipoDX
                lcCpt_Dx3 = ""
                lcDtipo3 = LxDx
                lcLab3 = " "
                
                oRsGrdOtrosCpt.Filter = "grupo=" & lnFor1 & " and subgrupo=" & lnFor2
                If oRsGrdOtrosCpt.RecordCount > 0 Then
                    lnIdTipo3 = sghActividadesTipo.TipoCPT
                    lcCpt_Dx3 = oRsGrdOtrosCpt!Codigo
                    lcDtipo3 = LxCPT
                Else
                    oRsDx.Filter = "grupo=" & lnFor1 & " and subgrupo=" & lnFor2
                    If oRsDx.RecordCount > 0 Then
                       lnIdTipo3 = sghActividadesTipo.TipoDX
                       lcCpt_Dx3 = oRsDx!CodigoCIE2004
                       lcDtipo3 = LxDx
                    Else
                       oRsSoloLabActividades.Filter = "grupo=" & lnFor1 & " and SubGrupo=" & lnFor2
                       If oRsSoloLabActividades.RecordCount > 0 Then
                          lnIdTipo3 = sghActividadesTipo.TipoLAB
                          lcCpt_Dx3 = sighEntidades.Lx_LabVacio
                          lcDtipo3 = sighEntidades.Lx_LabVacio
                          lcLab3 = oRsSoloLabActividades!lab
                       End If
                    End If
                End If
                oRsTmp1.AddNew
                oRsTmp1!Grupo = lnFor1
                oRsTmp1!subgrupoOrden = lnFor2
                oRsTmp1!EdadInicio = 0
                oRsTmp1!EdadFinal = 200
                oRsTmp1!PesoKgMenor = 0
                oRsTmp1!PesoKgMayor = 300
                oRsTmp1!idTipo = lnIdTipo3
                oRsTmp1!dTipo = lcDtipo3
                oRsTmp1!cpt_dx = lcCpt_Dx3
                oRsTmp1!lab = lcLab3
                oRsTmp1.Update
            Next
        Next
'        oRsTmp1.Sort = "idGrupo,subgrupoOrden"
        
        

        Dim oEdad As Edad, lbContiuar9 As Boolean
        '
       ' oEdad = calcularEdadDisgregada(mo_paciente.FechaNacimiento, mo_Atenciones.FechaIngreso)
        Select Case cmbTipoEdad1.ListIndex
        Case 0   'año
             oEdad.EdadAnio = Val(txtEdad1.Text)
        Case 1   'meses
             oEdad.EdadMes = Val(txtEdad1.Text)
        Case Else
             oEdad.EdadDia = Val(txtEdad1.Text)
        End Select
        oEdad.TipoEdad = cmbTipoEdad1.ListIndex + 1
        '
'        If oEdad.EdadAnio > 0 Then
'           oRsTmp1.Filter = "idtipoedad=1 and edadinicio" & IIf(ml_ups = "301202", "=", "<=") & oEdad.EdadAnio
'        ElseIf oEdad.EdadMes > 0 Then
'           oRsTmp1.Filter = "idtipoedad=2 and edadinicio=" & oEdad.EdadMes
'        Else
'          oRsTmp1.Filter = "idtipoedad=3"
'        End If
        If oRsTmp1.RecordCount > 0 Then
           Set oRsDx = Me.UcDiagnosticoDetalle1.DevuelveDx
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
              lnGrupo = oRsTmp1!Grupo

              lbPrimerReg = True
              Do While Not oRsTmp1.EOF And lnGrupo = oRsTmp1!Grupo
                 lbContiuar9 = True
                 If oEdad.EdadDia > 0 And oEdad.EdadAnio = 0 And oEdad.EdadMes = 0 Then
                    If Not (oEdad.EdadDia >= oRsTmp1!EdadInicio And oEdad.EdadDia <= oRsTmp1!EdadFinal) Then
                       lbContiuar9 = False
                    End If
                 End If

                 If (lnPesoKg >= oRsTmp1!PesoKgMenor And lnPesoKg <= oRsTmp1!PesoKgMayor) And lbContiuar9 = True Then
                    lnSubGrupo = oRsTmp1!subgrupoOrden
                    lcNombre = ""
                    Select Case oRsTmp1!idTipo
                    Case sghActividadesTipo.TipoCPT
                            Set oRsTmp2 = mo_AdminCaja.FactCatalogoServiciosSeleccionarPorCodigoOnombre(oRsTmp1!cpt_dx, "")
                            If oRsTmp2.RecordCount > 0 Then
                               lcNombre = Left(oRsTmp2!nombre, 255)
                            End If
                            oRsTmp2.Close
                    Case sghActividadesTipo.TipoLAB
                            'oRsSoloLabActividades.Filter = "grupo=" & lnGrupo & " and SubGrupo=" & lnSubGrupo
                            lcNombre = sighEntidades.Lx_LabVacio
                    Case sghActividadesTipo.TipoDX
                            Set oRsTmp2 = mo_AdminServiciosComunes.DiagnosticosSeleccionarXCodigo(oRsTmp1!cpt_dx)
                            If oRsTmp2.RecordCount > 0 Then
                               lcNombre = Left(oRsTmp2!descripcion, 255)
                            End If
                             oRsTmp2.Close
                    End Select
                    '
                    lcEligio = False
                    lcEligioLab = ""
                    lnEligioTipo = 102
                    lnEligioUPS = ml_idCuentaAtencion    'mo_Atenciones.IdServicioIngreso
                    ln_IdCuentaAtencion = ml_idCuentaAtencion: ln_IdOrden = 0: ln_Fua = 0: lc_Consultorio = ml_lcServicio
                    ln_idServicio = ml_idCuentaAtencion: lc_FuaCodigoPrestacion = "": ln_idServicioPaciente = mo_Atenciones.IdServicioIngreso
                    If oRsTmp1!idTipo = 1 Then
                        oRsGrdOtrosCpt.Filter = "grupo=" & lnGrupo & " and subgrupo=" & lnSubGrupo & _
                                                " and codigo='" & Trim(oRsTmp1!cpt_dx) & "'"
                        If oRsGrdOtrosCpt.RecordCount > 0 Then
                            lcEligio = True
                            lcEligioLab = oRsGrdOtrosCpt!labConfHIS
                            lnEligioTipo = IIf(IsNull(oRsGrdOtrosCpt!idTipoDiagnostico), 102, oRsGrdOtrosCpt!idTipoDiagnostico)
                            lnEligioUPS = oRsGrdOtrosCpt!idCuentaAtencion
                            ln_IdCuentaAtencion = oRsGrdOtrosCpt!idCuentaAtencion
                            ln_IdOrden = oRsGrdOtrosCpt!IdOrden
                            ln_Fua = oRsGrdOtrosCpt!FUA
                            lc_Consultorio = oRsGrdOtrosCpt!Consultorio
                            ln_idServicio = oRsGrdOtrosCpt!IdServicio
                            ln_idServicioPaciente = oRsGrdOtrosCpt!IdServicio
                        End If
                    Else
                        oRsDx.Filter = "grupo=" & lnGrupo & " and subgrupo=" & lnSubGrupo & _
                                       " and CodigoCIE2004='" & Trim(oRsTmp1!cpt_dx) & "'"
                        If oRsDx.RecordCount > 0 Then
                            lcEligio = True
                            lcEligioLab = IIf(IsNull(oRsDx!labConfHIS), "", oRsDx!labConfHIS)
                            lnEligioTipo = oRsDx!idTipoDiagnostico
                            lnEligioUPS = oRsDx!idCuentaAtencion
                            ln_IdCuentaAtencion = oRsDx!idCuentaAtencion
                            ln_Fua = oRsDx!FUA
                            lc_Consultorio = oRsDx!Consultorio
                            ln_idServicio = oRsDx!idCuentaAtencion
                            lc_FuaCodigoPrestacion = IIf(IsNull(oRsDx!FuaCodigoPrestacion), "", oRsDx!FuaCodigoPrestacion)
                            ln_idServicioPaciente = oRsDx!IdServicio
                        Else
                            oRsSoloLabActividades.Filter = "grupo=" & lnGrupo & " and SubGrupo=" & lnSubGrupo
                            If oRsSoloLabActividades.RecordCount > 0 Then
                               lcEligio = True
                               lcEligioLab = oRsSoloLabActividades!lab
                            End If
                        End If

                    End If
                    '
                    oRsActividades.AddNew
                    If lbPrimerReg = True Then
                       lbPrimerReg = False
                       oRsActividades!GrupoTIT = Trim(Str(lnGrupo))
                    Else
                       oRsActividades!GrupoTIT = ""
                    End If
                    oRsActividades!Grupo = lnGrupo
                    oRsActividades!SubGrupo = lnSubGrupo
                    oRsActividades!lab = IIf(IsNull(oRsTmp1!lab), " ", oRsTmp1!lab)
                    oRsActividades!ID = oRsTmp1!cpt_dx
                    oRsActividades!tipo = oRsTmp1!dTipo
                    oRsActividades!nombre = lcNombre
                    oRsActividades!elija = lcEligio
                    oRsActividades!elijaTipo = lnEligioTipo - 100
                    oRsActividades!ElijaUPS = lnEligioUPS
                    oRsActividades!ElijaLab = lcEligioLab
                    oRsActividades!idCuentaAtencion = ln_IdCuentaAtencion
                    oRsActividades!IdOrden = ln_IdOrden
                    oRsActividades!FUA = ln_Fua
                    oRsActividades!Consultorio = lc_Consultorio
                    oRsActividades!IdServicio = ln_idServicio
                    oRsActividades!FuaCodigoPrestacion = lc_FuaCodigoPrestacion
                    oRsActividades!idTipo = oRsTmp1!idTipo
                    oRsActividades!idServicioPaciente = ln_idServicioPaciente
                    oRsActividades.Update
                 End If
                 oRsTmp1.MoveNext
                 If oRsTmp1.EOF Then
                    Exit Do
                 End If
              Loop
           Loop

           oRsGrdOtrosCpt.Filter = ""
           If oRsGrdOtrosCpt.RecordCount > 0 Then
              oRsGrdOtrosCpt.MoveFirst
           End If
           oRsDx.Filter = ""
           If oRsDx.RecordCount > 0 Then
              oRsDx.MoveFirst
           End If
        End If
        oRsTmp1.Close
        If oRsActividades.RecordCount > 0 Then
            Dim oAdmisionCEatencSimultanea As New AdmisionCEprogramas
            Dim oRsItemsElegidos As New Recordset
            oAdmisionCEatencSimultanea.FormLlamante = "ACTIVIDADES"
            Set oAdmisionCEatencSimultanea.oRsFua = oRsActividades
            Set oAdmisionCEatencSimultanea.oRsItemsElegidos = oRsTipoDx
            oAdmisionCEatencSimultanea.Show 1
            If oAdmisionCEatencSimultanea.idCuentaAtencion = 1 Then
               Set oRsItemsElegidos = oAdmisionCEatencSimultanea.ItemsMasivosElegidos
               '
               oRsSoloLabActividades.Filter = ""
               CreaYllenaTemporalesActividades True
               oRsItemsElegidos.Filter = "idTipo=" & sghActividadesTipo.TipoLAB
               If oRsItemsElegidos.RecordCount > 0 Then
                  oRsItemsElegidos.MoveFirst
                  Do While Not oRsItemsElegidos.EOF
                     oRsSoloLabActividades.AddNew
                     oRsSoloLabActividades!Grupo = oRsItemsElegidos!Grupo
                     oRsSoloLabActividades!SubGrupo = oRsItemsElegidos!SubGrupo
                     oRsSoloLabActividades!lab = oRsItemsElegidos!ElijaLab
                     oRsSoloLabActividades.Update
                     oRsItemsElegidos.MoveNext
                  Loop
               End If
               '
               ActividadesHIS oRsItemsElegidos
            End If
            Set oAdmisionCEatencSimultanea = Nothing
            Set oRsItemsElegidos = Nothing
            
         End If
        Set oRsActividades = Nothing
        Set oRsTmp1 = Nothing
        Set oRsDx = Nothing
        Set oFactOrdenServicio = Nothing
        Set oDOFactOrdenServicio = Nothing
        Set mrs_FacturacionProductos = Nothing
        Set oDoCatalogoServicioHosp = Nothing

End Sub


Function ValidarDatosObligatoriosCS() As Boolean
   ValidarDatosObligatoriosCS = False
   ms_MensajeError = ""
    Set mRs_Productos = grdServicios.DevuelveProductos
   If ms_MensajeError <> "" Then
       MsgBox ms_MensajeError, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatoriosCS = True
End Function

Private Sub btnguardarserv_Click()
If ValidarDatosObligatoriosCS() Then
If AgregarDatosCS() Then
                MsgBox "Se registrarón correctamente los datos " + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
                grdServicios.LimpiarGrilla
            Else
                MsgBox "No se registrarón los datos " + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
            End If
       End If
End Sub
Function AgregarDatosCS() As Boolean
    Dim lbAgregarDatosCS As Boolean   '
    lbAgregarDatosCS = mo_ReglasFarmacia.usp_catalogoserviciosagregar(mRs_Productos)
    ms_MensajeError = mo_ReglasFarmacia.MensajeError
    AgregarDatosCS = lbAgregarDatosCS
End Function

Private Sub Form_Initialize()
    Set mo_cmbIdDestinoAtencion.MiComboBox = cmbIdDestinoAtencion
    Set mo_cmbIdTipoReferenciaDestino.MiComboBox = cmbIdTipoReferenciaDestino
    Set mo_cmbIdCondicionEnElServicio.MiComboBox = cmbIdCondicionEnElServicio
    Set mo_cmbIdCondicionEnElEstablecimiento.MiComboBox = cmbIdCondicionEnElEstablecimiento
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub

'Actualizado 09102014
Private Sub grdOtrosCpt_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Cancel = True
End Sub

'debb-27/05/2015
Private Sub grdOtrosCpt_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    

        grdOtrosCpt.Font.Name = "Arial Narrow"
        grdOtrosCpt.Font.Size = 8
        grdOtrosCpt.Bands(0).Columns("IdPuntoCarga").Hidden = True
        grdOtrosCpt.Bands(0).Columns("IdProducto").Hidden = True
        
        grdOtrosCpt.Bands(0).Columns("IdCuentaAtencion").Hidden = True
        grdOtrosCpt.Bands(0).Columns("IdOrden").Hidden = True
        
        grdOtrosCpt.Bands(0).Columns("consultorio").Width = 1700
        grdOtrosCpt.Bands(0).Columns("consultorio").Header.Appearance.ForeColor = vbWhite
        grdOtrosCpt.Bands(0).Columns("consultorio").Header.Appearance.BackColor = vbRed
        grdOtrosCpt.Bands(0).Columns("consultorio").Header.Appearance.Font.Bold = True
        grdOtrosCpt.Bands(0).Columns("consultorio").Header.Caption = "UPS"
       
        
        grdOtrosCpt.Bands(0).Columns("codigo").Width = 800
        grdOtrosCpt.Bands(0).Columns("Nombre").Width = 5000
        grdOtrosCpt.Bands(0).Columns("cantidad").Width = 700
        grdOtrosCpt.Bands(0).Columns("cantidad").Hidden = True
        grdOtrosCpt.Bands(0).Columns("precio").Width = 600
        grdOtrosCpt.Bands(0).Columns("precio").Hidden = True
        grdOtrosCpt.Bands(0).Columns("total").Width = 900
        grdOtrosCpt.Bands(0).Columns("total").Hidden = True
        
        grdOtrosCpt.Bands(0).Columns("labConfHIS").Header.Caption = "Lab"    '"Detalle adicional"
        grdOtrosCpt.Bands(0).Columns("labConfHIS").Width = 2000
        grdOtrosCpt.Bands(0).Columns("grupo").Width = 300
        grdOtrosCpt.Bands(0).Columns("subgrupo").Width = 300
        grdOtrosCpt.Bands(0).Columns("idServicio").Hidden = True
        grdOtrosCpt.Bands(0).Columns("idtipoDiagnostico").Hidden = True
        grdOtrosCpt.Bands(0).Columns("ups").Hidden = True
 
        Dim oRsLab As New Recordset
        Set oRsLab = mo_AdminServiciosComunes.DevuelveHIS_SITUACIOporDescripcion()
        With grdOtrosCpt.ValueLists.Add("LabLista").ValueListItems
           oRsLab.MoveFirst
           Do While Not oRsLab.EOF
              .Add Trim(oRsLab!valores), oRsLab!valores
              oRsLab.MoveNext
           Loop
        End With
        oRsLab.Close
        Set oRsLab = Nothing
        grdOtrosCpt.Bands(0).Columns("labConfHIS").ValueList = "LabLista"
        grdOtrosCpt.Bands(0).Columns("labConfHIS").ButtonDisplayStyle = ssButtonDisplayStyleAlways
         

End Sub

Private Sub cmbIdDestinoAtencion_Click()
    lbHuboCambioEnDato = True
    Dim sCodigoDestino As String
    If cmbIdDestinoAtencion.Text = "" Then
       Exit Sub
    End If
    
    sCodigoDestino = Trim(Split(cmbIdDestinoAtencion.Text, " = ")(0))
    ucCitasLista1.Visible = False       'franklin 2017
    If sCodigoDestino <> "R" And sCodigoDestino <> "C" Then
        mo_cmbIdTipoReferenciaDestino.BoundText = ""
        Me.txtIdEstablecimientoDestino.Tag = ""
        Me.txtIdEstablecimientoDestino = ""
        txtNombreDestinoReferencia.Text = ""
        txtNroReferenciaDestino.Text = ""
        'debb-21/06/2016 (inicio)
        cmbServicioReferenciaD.Text = ""
        txtFextension.Text = sighEntidades.FECHA_VACIA_DMY
        txtFtramite.Text = sighEntidades.FECHA_VACIA_DMY
        'debb-21/06/2016 (fin)
    End If
    
    HabilitarFrameDestino False
    Select Case sCodigoDestino
'    Case "D", "H", "O", "M"
'        HabilitarFrameDestino False
    Case "R"
        HabilitarFrameDestino True
        Me.fraDatosReferenciaDestino = "Referencia destino "
        Me.lblIdTipoReferenciaDestino = "Tipo Referencia"
        Me.lblIdEstablecimientoDestino = "Estab. Referencia"
        mo_cmbIdTipoReferenciaDestino.BoundText = "1"
    Case "C"
        HabilitarFrameDestino True
        Me.fraDatosReferenciaDestino = "Contrareferencia destino "
        Me.lblIdTipoReferenciaDestino = "Tipo Contrarefer."
        Me.lblIdEstablecimientoDestino = "Estab. Contrarefer."
        mo_cmbIdTipoReferenciaDestino.BoundText = "1"
    Case "I"
        ucCitasLista1.Visible = True    'franklin 2017
    End Select
    
   If sCodigoDestino = "R" Or sCodigoDestino = "C" Then
        mo_cmbIdTipoReferenciaDestino.BoundText = "1"
        txtFextension.Text = Format(mo_Atenciones.FechaIngreso, sighEntidades.DevuelveFechaSoloFormato_DMY)
        txtFtramite.Text = Format(mo_Atenciones.FechaIngreso, sighEntidades.DevuelveFechaSoloFormato_DMY)
        txtNroReferenciaDestino.Text = mo_AdminServiciosComunes.CalculaNUMEROREFERENCIA(IIf(sCodigoDestino = "C", True, False))
        
        txtIdEstablecimientoDestino.Tag = ""
        txtIdEstablecimientoDestino.Text = ""
        txtNombreDestinoReferencia = ""
        If mo_DoAtencionDatosAdicionales.IdEstablecimientoOrigen > 0 And sCodigoDestino = "C" Then
            Dim oDoEstablecimiento As New DOEstablecimiento
            Set oDoEstablecimiento = mo_AdminServiciosComunes.EstablecimientosSeleccionarPorId(mo_DoAtencionDatosAdicionales.IdEstablecimientoOrigen)
            If Not oDoEstablecimiento Is Nothing Then
                txtIdEstablecimientoDestino.Tag = oDoEstablecimiento.IdEstablecimiento
                txtIdEstablecimientoDestino.Text = oDoEstablecimiento.Codigo
                txtNombreDestinoReferencia = oDoEstablecimiento.nombre
            End If
            Set oDoEstablecimiento = Nothing
        End If
    End If
    
End Sub

Private Sub cmbIdDestinoAtencion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdDestinoAtencion
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdDestinoAtencion_LostFocus()
    If lbHuboCambioEnDato = True Then
      sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, cmbIdDestinoAtencion.Text
      lbHuboCambioEnDato = False
    End If

    Dim oDOTipoDestinoAtencion As New DOTipoDestinoAtencion

    If cmbIdDestinoAtencion.Text <> "" Then
         Set oDOTipoDestinoAtencion = mo_AdminAdmision.TiposDestinoAtencionSeleccionarPorCodigo(Trim(Split(cmbIdDestinoAtencion.Text, " = ")(0)), ml_TipoServicio)
         If oDOTipoDestinoAtencion.IdDestinoAtencion <> 0 Then
             mo_cmbIdDestinoAtencion.BoundText = oDOTipoDestinoAtencion.IdDestinoAtencion
        End If
    End If
    mo_Formulario.MarcarComoVacio cmbIdDestinoAtencion
    Set oDOTipoDestinoAtencion = Nothing
End Sub

Private Sub cmbIdDestinoAtencion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
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










Private Sub TabAtencion_Click(PreviousTab As Integer)
    If TabAtencion.Tab = 2 Then    'se ingresa RECETA,ORDENES MEDICAS
       If Me.UcProgramaMaterno.Visible = True Then
          If wxParametro545 = "S" Then
             UcRecetas1.ActualizaDxEnGrilla Me.UcProgramaMaterno.DevuelveProDiagnosticos
          End If
       ElseIf Me.ucPerinatalAS1.Visible = True Then
          If wxParametro545 = "S" Then
             UcRecetas1.ActualizaDxEnGrilla Me.ucPerinatalAS1.DevuelveDxMorbilidad
          End If
       ElseIf Me.ucPerinatal1.Visible = True Then
          If wxParametro545 = "S" Then
             UcRecetas1.ActualizaDxEnGrilla Me.ucPerinatal1.DevuelveDxMorbilidad
          End If
       Else
          UcRecetas1.ActualizaDxEnGrilla UcDiagnosticoDetalle1.DevuelveDx
       End If
    End If
End Sub



Private Sub txtantecedAlergico_LostFocus()
    If oDoPacienteDatosAdd.antecedAlergico <> "" And Me.txtantecedAlergico.Text = "" Then
       If EliminaAntecedentePersonal("Eliminó Antecente Alégico") = False Then
          Me.txtantecedAlergico.Text = oDoPacienteDatosAdd.antecedAlergico
       End If
    End If

End Sub




Private Sub txtAntecedentes_LostFocus()
    If oDoPacienteDatosAdd.antecedentes <> "" And Me.txtAntecedentes.Text = "" Then
       If EliminaAntecedentePersonal("Eliminó otros Antecedentes") = False Then
          Me.txtAntecedentes.Text = oDoPacienteDatosAdd.antecedentes
       End If
    End If

End Sub

Private Sub txtantecedFamiliar_LostFocus()
    If oDoPacienteDatosAdd.antecedFamiliar <> "" And Me.txtantecedFamiliar.Text = "" Then
       If EliminaAntecedentePersonal("Eliminó Antecente Familiar") = False Then
          Me.txtantecedFamiliar.Text = oDoPacienteDatosAdd.antecedFamiliar
       End If
    End If

End Sub



Private Sub txtantecedObstetrico_LostFocus()
    If oDoPacienteDatosAdd.antecedObstetrico <> "" And Me.txtantecedObstetrico.Text = "" Then
       If EliminaAntecedentePersonal("Eliminó Antecente Obstétrico") = False Then
          Me.txtantecedObstetrico.Text = oDoPacienteDatosAdd.antecedObstetrico
       End If
    End If

End Sub

Private Sub txtantecedPatologico_LostFocus()
    If oDoPacienteDatosAdd.antecedPatologico <> "" And Me.txtantecedPatologico.Text = "" Then
       If EliminaAntecedentePersonal("Eliminó Antecente Patológico") = False Then
          Me.txtantecedPatologico.Text = oDoPacienteDatosAdd.antecedPatologico
       End If
    End If

End Sub

Private Sub txtantecedQuirurgico_LostFocus()
    If oDoPacienteDatosAdd.antecedQuirurgico <> "" And Me.txtantecedQuirurgico.Text = "" Then
       If EliminaAntecedentePersonal("Eliminó Antecente Quirúrgico") = False Then
          txtantecedQuirurgico.Text = oDoPacienteDatosAdd.antecedQuirurgico
       End If
    End If
End Sub



Private Sub txtCitaAntecedente_Change()
lbHuboCambioEnDato = True
End Sub

Private Sub txtCitaAntecedente_LostFocus()
    If lbHuboCambioEnDato = True Then
      sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, txtCitaAntecedente.Text
      lbHuboCambioEnDato = False
    End If
End Sub

Private Sub txtCitaDxMedico_Change()
lbHuboCambioEnDato = True
End Sub

Private Sub txtCitaExamenClinico_Change()
lbHuboCambioEnDato = True
End Sub

Private Sub txtCitaExamenClinico_LostFocus()
    If lbHuboCambioEnDato = True Then
      sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, txtCitaExamenClinico.Text
      lbHuboCambioEnDato = False
    End If
End Sub

Private Sub txtCitaMotivo_Change()
lbHuboCambioEnDato = True
End Sub

Private Sub txtCitaMotivo_LostFocus()
    If lbHuboCambioEnDato = True Then
      sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, txtCitaMotivo.Text
      lbHuboCambioEnDato = False
    End If
End Sub

Private Sub TxtCitaTratamiento_Change()
lbHuboCambioEnDato = True
End Sub

Private Sub TxtCitaTratamiento_LostFocus()
    If lbHuboCambioEnDato = True Then
      sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, TxtCitaTratamiento.Text
      lbHuboCambioEnDato = False
    End If
    If Trim(TxtCitaTratamiento.Text) <> "" Then
        Me.UcRecetas1.Tratamiento = Trim(TxtCitaTratamiento.Text)
    End If
End Sub



Private Sub cmbIdTipoReferenciaDestino_Click()

    txtIdEstablecimientoDestino.Tag = ""
    txtIdEstablecimientoDestino = ""
    txtNombreDestinoReferencia = ""
    txtNroReferenciaDestino.Text = ""
    
End Sub

Private Sub cmbIdTipoReferenciaDestino_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoReferenciaDestino
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdTipoReferenciaDestino_LostFocus()
   If cmbIdTipoReferenciaDestino.Text <> "" Then
       mo_cmbIdTipoReferenciaDestino.BoundText = Val(Split(cmbIdTipoReferenciaDestino.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdTipoReferenciaDestino
End Sub

Private Sub cmbIdTipoReferenciaDestino_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub



Private Sub btnBuscarEstablecimientoDestino_Click()
    If cmbIdTipoReferenciaDestino.Text <> "" Then
       CompletarDatosDeEstablecimiento txtIdEstablecimientoDestino, txtNombreDestinoReferencia, Val(mo_cmbIdTipoReferenciaDestino.BoundText)
    End If
End Sub

Sub CompletarDatosDeEstablecimiento(txtIdEstablecimiento As TextBox, lblNombreEstablecimiento As TextBox, lTipoReferencia As Long)
    
    If lTipoReferencia = 1 Then
        Dim oBusqueda As New SIGHNegocios.BuscaEstablecimientos
        Dim oDoEstablecimiento As New DOEstablecimiento
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
        Dim oBusquedaNM As New SIGHNegocios.BuscaEstablecNoMinsa
        Dim oDoEstablecimientoNM As New DOEstablecimientoNoMinsa
        oBusquedaNM.lcNombrePc = mo_lcNombrePc
        oBusquedaNM.idUsuario = ml_idUsuario
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









Private Sub txtHemoglobina_Change()
lbHuboCambioEnDato = True
End Sub

Private Sub txtHemoglobina_LostFocus()
    If lbHuboCambioEnDato = True Then
      sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, txtHemoglobina.Text
      lbHuboCambioEnDato = False
    End If
End Sub

Private Sub txtIdEstablecimientoDestino_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdEstablecimientoDestino
    If KeyCode = vbKeyF1 Then
        btnBuscarEstablecimientoDestino_Click
    End If
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdEstablecimientoDestino_LostFocus()
    CompletarDatosDelEstablecimientoEnElLostFocus txtIdEstablecimientoDestino, txtNombreDestinoReferencia, Val(mo_cmbIdTipoReferenciaDestino.BoundText)
    mo_Formulario.MarcarComoVacio txtIdEstablecimientoDestino
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


Private Sub txtIdEstablecimientoDestino_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub













Sub CargarDatosAlFormulario()

    
    HabilitarFrameDestino False
    If mi_Opcion = sghModificar Then
        mo_Formulario.HabilitarDeshabilitar Me.cmbIdCondicionEnElEstablecimiento, True
        mo_Formulario.HabilitarDeshabilitar Me.cmbIdCondicionEnElServicio, True
    Else
        mo_Formulario.HabilitarDeshabilitar Me.cmbIdCondicionEnElEstablecimiento, False
        mo_Formulario.HabilitarDeshabilitar Me.cmbIdCondicionEnElServicio, False
    End If
    
    
    mo_Formulario.HabilitarDeshabilitar txtNombreDestinoReferencia, False
    mo_Formulario.HabilitarDeshabilitar txtIdEstablecimientoDestino, False
    
    
    
    
    Select Case mi_Opcion
     Case sghAgregar
     Case sghModificar
         CargarDatosAlosControles
     Case sghConsultar
         CargarDatosAlosControles
     Case sghEliminar
         CargarDatosAlosControles
    End Select
    
    Select Case mi_Opcion
     Case sghAgregar
     Case sghModificar
     Case sghConsultar
        DeshabilitarControlesParaEdicion
        Me.btnAceptar.Visible = False
    Case sghEliminar
        DeshabilitarControlesParaEdicion
    End Select
 
End Sub
Sub DeshabilitarControlesParaEdicion()
    
    
    HabilitarFrameDestino False
   

End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla CuentasAtencion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------



Sub Form_Load()
    If mo_lbCargaTablasUnaVez = True Then
        mo_Formulario.HabilitarDeshabilitar lblMedico, False
        lbTieneLicenciaParaMensajeAcelulares = mo_sighProxies.VerificaLicenciaMensajeTexto
        lbCargaTablasUnaVez = False
        VerificaPermisos
        InicilizarParametros
        Me.UcDiagnosticoDetalle1.Inicializar
        'franklin 2017
        Me.ucCitasLista1.Inicializar
        If wxParametro518 <> "S" Then Me.ucCitasLista1.InhabilitaDiario
        '
        '
        UcRecetas1.Inicializar
        UcRecetas1.DatoCabeceraReceta = "(N° Cuenta=" & Trim(Str(ml_idCuentaAtencion)) & ") Paciente=" & lcHistoriaYpaciente
        UcRecetas1.Height = 9705
        UcRecetas1.lnWnd = Me.hwnd
        CargarComboBoxes
        '
        '
        lbBuscaDNIenReniec = IIf(wxParametro296 = "S", True, False)
        If lbBuscaDNIenReniec = True Then
           mo_Reniec.SeAccesaAlaWebDesdeGalenhos = True
           mo_Reniec.Inicializar
        End If
        '
        Set oRsTipoDx = mo_AdminServiciosComunes.SubclasificacionDiagnosticosSeleccionarDxConsultaExterna
        '
        mo_Apariencia.ConfigurarFilasBiColores Me.grdOtrosCpt, sighEntidades.GrillaConFilasBicolor
    End If
    '
    SiempreCargaPorMovimiento
End Sub

Sub EligeTabAtencion()
   TabAtencion.Tab = Val(wxParametro274)
End Sub

Sub HabilitaModulosPerinatalYmaterno()
        Dim lbElConsultorioUsaFUA As Boolean
        TabDx.TabVisible(1) = False
        TabDx.TabVisible(2) = False
        
        
        ConfigurarGrdServicios
        
        
        mo_AdminArchivoClinico.ServicioSeUsanModulosPerinatalyMaterno ml_IdCita, lbElConsultorioUsaModuloPerinatal, _
                                                              lbElConsultorioUsaModuloMaterno, lbElConsultorioUsaFUA
        'Modulo niño sano - habilita
        If lbElConsultorioUsaModuloPerinatal = True And _
                            mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE Then
           'debb-09/06/2016 (inicio)
           If wxParametro502 = "S" Then
                Me.ucPerinatalAS1.Visible = True
                Me.ucPerinatalAS1.Height = 7185: Me.ucPerinatalAS1.Left = 45: Me.ucPerinatalAS1.Top = 405: Me.ucPerinatalAS1.Width = 11655
                TabAtencion.Tab = 1
                TabDx.TabVisible(1) = True
                TabDx.Tab = 1
           Else
                Me.ucPerinatal1.Visible = True
                Me.ucPerinatal1.Height = 7185: Me.ucPerinatal1.Left = 45: Me.ucPerinatal1.Top = 405: Me.ucPerinatal1.Width = 11655
                btnAntecedentesPersonales.Visible = True
                cmdGenerarPlanAtencion.Visible = True
                TabAtencion.Tab = 1
                TabDx.TabVisible(1) = True
                TabDx.Tab = 1
                Me.ucPerinatal1.EnfoqueTabProcedimientos
           End If
'            Me.ucPerinatal1.Visible = True
'            Me.ucPerinatal1.Height = 7185: Me.ucPerinatal1.Left = 45: Me.ucPerinatal1.Top = 405: Me.ucPerinatal1.Width = 11655
'            btnAntecedentesPersonales.Visible = True
'            cmdGenerarPlanAtencion.Visible = True
'            TabAtencion.Tab = 1
'            TabDx.TabVisible(1) = True
'            TabDx.Tab = 1
'            Me.ucPerinatal1.EnfoqueTabProcedimientos
           'debb-09/06/2016 (fin)
        Else
           Me.ucPerinatal1.Visible = False
           Me.ucPerinatalAS1.Visible = False
           'mgaray201410f
           TabDx.TabVisible(0) = True
           btnAntecedentesPersonales.Visible = False
           cmdGenerarPlanAtencion.Visible = False
            TabDx.Tab = 0
           TabDx.TabCaption(0) = "3.2.1 Diagnósticos y CPT" 'Actualizado 19092014
           TabDx.TabVisible(1) = False
        End If
'AGREGADO POR FRANK - MODULO PROGRAMA MATERNO
        'Modulo materno - habilita
        If lbElConsultorioUsaModuloMaterno = True And _
                        mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE Then
           Me.UcProgramaMaterno.Visible = True
           TabAtencion.Tab = 2
           'TabAtencion.Caption = "3.3 Módulo Materno"
           TabDx.TabVisible(2) = True
           TabDx.Tab = 2
        Else
           Me.UcProgramaMaterno.Visible = False
           'mgaray201410f
           TabDx.TabVisible(0) = True
           If Me.UcProgramaMaterno.Visible = False Then
                TabDx.Tab = 0
                TabDx.TabCaption(0) = "3.2.1 Diagnósticos y CPT" 'Actualizado 19092014
                TabDx.TabVisible(2) = False
           End If
        End If
        '
        btnImprimeFichaSIS.Visible = False
        If lbElConsultorioUsaFUA = False Then
           wxParametro302 = "N"
        Else
           wxParametro302 = lcBuscaParametro.SeleccionaFilaParametro(302)
           If wxParametro302 = "S" Then
              btnImprimeFichaSIS.Visible = True
           End If
        End If
        
        
        '
End Sub


Sub ConfigurarGrdServicios()
    'grdServicios.movNumero = ml_movNumero
    'grdServicios.IdAlmacen = 0
    'grdServicios.FechaMinimaDespacho = CDate(lcBuscaParametro.RetornaFechaServidorSQL) + Val(lcBuscaParametro.SeleccionaFilaParametro(220))
    grdServicios.Inicializar
End Sub


Sub SiempreCargaPorMovimiento()
    If mo_lbNuevoMovimiento = True Then
        sighEntidades.ParaAuditoria = ""
        lbImpresionDeAtencionDistintoAlGrabar = True
        ucCitasLista1.Visible = False    'franklin 2017
        
        HabilitaModulosPerinatalYmaterno
        '
        'TabAtencion.Height = 6065  ' 6465
        TabDx.Height = 5505
        Me.ucPerinatal1.Height = 1
        Me.ucPerinatalAS1.Height = 1            'debb-09/06/2016
'AGREGADO POR FRANK - MODULO PROGRAMA MATERNO
        Me.UcProgramaMaterno.Height = 1
        If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE Then
         '  TabAtencion.Height = TabAtencion.Height + 1930
           TabDx.Height = TabDx.Height + 1930
           Me.ucPerinatal1.Height = 9705
           Me.ucPerinatalAS1.Height = 9705          'debb-09/06/2016
'AGREGADO POR FRANK - MODULO PROGRAMA MATERNO
           Me.UcProgramaMaterno.Height = 9705
           '
        End If
        '
        mo_lbNuevoMovimiento = False
        lbYaSeTransfirioHCdeUnServicioAotro = False
        mb_FormLoading = True
        btnImprimeFichaSIS.Visible = False
'        lblAlertaTemperatura.Visible = False
        
        Select Case mi_Opcion
        Case sghAgregar
            Me.Caption = "Agrega Admisión de CE"
        Case sghModificar
            Me.Caption = "Modifica Admisión de CE"
            EligeTabAtencion
            
            'grdServicios.AgregaRegistro
            
            
        Case sghConsultar
            Me.Caption = "Consulta Admisión de CE"
        Case sghEliminar
            Me.Caption = "Elimina Admisión de CE"
        End Select
        '
        '
        LimpiaTodosControles
        CargarDatosAlFormulario
        '
        mo_Formulario.ConfigurarTipoLetra "Tahoma", "11", Me
        If mi_Opcion = sghAgregar Then
           btnAceptar.Enabled = True
        End If
        btnAceptar.Visible = True
        '
        lc_AntecedentePersonal = ""
        '
        lbElMedicoNOregistraDatosCE = mo_AdminServiciosComunes.lbElMedicoNOregistraDatosCE(ml_IdServicio)
        '
        
    End If
End Sub

Sub LimpiaTodosControles()
    If mi_Opcion = sghAgregar Then
            mo_Pacientes.idPaciente = 0
            '
            '
            Me.idCuentaAtencion = 0
            Me.idPaciente = 0
            mo_cmbIdDestinoAtencion.BoundText = ""
            mo_cmbIdTipoReferenciaDestino.BoundText = ""
            Me.txtIdEstablecimientoDestino.Tag = ""
            txtNroReferenciaDestino.Text = ""
            cmbServicioReferenciaD.Text = ""            'debb-21/06/2016
            txtFextension.Text = sighEntidades.FECHA_VACIA_DMY  'debb-21/06/2016
            txtFtramite.Text = sighEntidades.FECHA_VACIA_DMY    'debb-21/06/2016
            
            '
            lnIdDistritoSIS = 0: lnIdSexoSIS = 0: ldFechaNacimientoSIS = 0: lcSnombreSIS = "": lnIdPlanSIS = 0
            '
    End If
    '
    lcHistoriaYpaciente = ""
    txtCitaMotivo.Text = ""
    txtCitaExamenClinico.Text = ""
    txtCitaExClinicos.Text = ""
    Me.TxtCitaTratamiento.Text = ""
    Me.txtCitaObservaciones.Text = ""
'    Me.txtPeso.Text = ""
'    Me.txtPresion.Text = "___/___"
'    Me.txtTemperatura.Text = ""
'    Me.txtTalla.Text = ""
'    Me.txtPulso.Text = ""
'    Me.txtFrespiratoria.Text = ""
    txtCitaDxMedico.Text = ""
    Me.txtProximaCita.Text = ""
    lblProximaCita.Caption = ""
    txtNroHijos.Text = ""
    Me.txtAntecedentes.Text = ""
    Me.txtantecedAlergico.Text = ""
    Me.txtantecedFamiliar.Text = ""
    Me.txtantecedObstetrico.Text = ""
    Me.txtantecedPatologico.Text = ""
    Me.txtantecedQuirurgico.Text = ""
    Me.txtCitaAntecedente.Text = ""
    '
    Me.UcDiagnosticoDetalle1.GenerarRecordsetTemporal      ' Me.UcDiagnosticoDetalle1.LimpiarDatos
End Sub


'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla CuentasAtencion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Activate()
   SiempreCargaPorMovimiento
   UcRecetas1.Height = 9705
   btnAceptar.Enabled = True
   ConfigurarTabTratamiento 'frank 30092015
   If mi_Opcion <> sghAgregar Then
       If Not mb_ExistenDatos Then
            Me.Visible = False
            
            LimpiarVariablesDeMemoria
       End If
       
       If ml_ldFechaIngreso < ldFechaActualServidor And (mi_Opcion = sghModificar Or mi_Opcion = sghEliminar) Then
          If lbTienePermisoParaRegistrarAtencionesPasadas = False Then
            MsgBox "No puede Modificar/Eliminar ATENCION DE CITAS de días menores a: " & ldFechaActualServidor, vbInformation, Me.Caption
            Me.Visible = False
            LimpiarVariablesDeMemoria
          End If
       End If
       If mo_Atenciones.HoraEgreso = "" And (mi_Opcion = sghConsultar Or mi_Opcion = sghEliminar) Then
           MsgBox "No hay nada registrado para CONSULTAR o ELIMINAR", vbInformation, Me.Caption
           Me.Visible = False
           LimpiarVariablesDeMemoria
       End If
   Else
       btnBuscaHistoricos.Visible = False
       If ml_ldFechaIngreso < ldFechaActualServidor Then
            MsgBox "No puede registrar CITAS de días menores a: " & ldFechaActualServidor, vbInformation, Me.Caption
            Me.Visible = False
            LimpiarVariablesDeMemoria
        End If
   End If
   If mb_FormLoading Then
        On Error Resume Next
        Select Case mi_Opcion
        Case sghAgregar
        Case sghModificar
        
        
        grdServicios.CargaProductosPorIdAtencion
        
            If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE Then
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
    If mi_Opcion = sghConsultar Then
        btnAceptar.Enabled = False
    End If
End Sub

Sub ConfigurarTabTratamiento()
    If lcBuscaParametro.SeleccionaFilaParametro(362) = "N" Then
        TabAtencion.TabVisible(3) = False
        TabAtencion.TabCaption(4) = "3.4 Destino Atención"
    ElseIf lcBuscaParametro.SeleccionaFilaParametro(362) = "S" Then
        TabAtencion.TabVisible(3) = True
        TabAtencion.TabCaption(4) = "3.5 Destino Atención"
    End If
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
    
    Select Case KeyCode
    'Case vbKeyEscape
    '    btnCancelar_Click
    Case vbKeyF2
        btnAceptar_Click
    Case vbKeyF12
         On Error Resume Next
         UcDiagnosticoDetalle1.IdListBarItem = mo_lnIdTablaLISTBARITEMS
         UcDiagnosticoDetalle1.FocusEnDx
         
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
   Select Case mi_Opcion
   Case sghAgregar
   Case sghModificar, sghEliminar
   
   
        
       
       If ValidarDatosObligatorios() Then
            CargaDatosAlObjetosDeDatos
            If ValidarReglas() Then
                'mgaray20141024
                ml_idOrdenServicioInmunizaciones = getIdOrdenServicioInmunizaciones(mo_Atenciones.idAtencion)
               If ModificarDatos() Then
                   btnImprimeAtencion.Enabled = True
                   
                   ActualizaVacunasDesdeModuloPerinatal
                   ActualizaVacunasDesdeModuloPerinatalAS       'debb-09/06/2016
                   '
                   '
                   
                   If lbTieneLicenciaParaMensajeAcelulares = True Then
                        Dim oMensajeCelular As New SIGHProxies.Procesos
                        oMensajeCelular.MensajeCelularEnviarSegunDxGRAVE mo_Pacientes, Me.UcDiagnosticoDetalle1.DevuelveDx, "ATENCIONCE", _
                                                                          mo_Atenciones.idCuentaAtencion
                        Set oMensajeCelular = Nothing
                   End If
                   
                   
                   If Me.ucPerinatal1.Visible = True Then
                        Call cargaCptDesdeProgramasPerinatalMaterno(True, Me.ucPerinatal1.DevuelveCptFrecuentes, ml_idOrdenServicioInmunizaciones)
                   End If
                   If Me.UcProgramaMaterno.Visible = True Then
                        Call cargaCptDesdeProgramasPerinatalMaterno(False, Me.UcProgramaMaterno.DevuelveProProcedimientos, 0)
                   End If
                   
                   ms_NombrePaciente = mo_paciente.ApellidoPaterno + " " + mo_paciente.ApellidoMaterno + " " + mo_paciente.PrimerNombre
                   MsgBox " Los datos se modificaron correctamente, para la Cuenta N°: " & Trim(Str(ml_idCuentaAtencion)) & DevuelveNroRecetasGeneradas, vbInformation, Me.Caption
                   
                   
                   If ValidarDatosObligatoriosCS() Then
                If AgregarDatosCS() Then
                'MsgBox "Se registrarón correctamente los datos " + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
                grdServicios.LimpiarGrilla
            'Else
                'MsgBox "No se registrarón los datos " + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
            End If
            End If
                   
                   
                   
                   
                   
                   If wxParametro302 = "S" And _
                      mo_Atenciones.IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS And _
                      lbElMedicoNOregistraDatosCE <> "S" Then
                         btnImprimeFichaSIS_Click
                   End If
                   'El formulario atenciones no debe cerrarse para los pacientes SIS, requerimiento Hosp. Socorro de Ica
                   If Not (wxParametro302 = "S" And mo_Atenciones.IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS And _
                                                                                            lbElMedicoNOregistraDatosCE <> "S") Then
                        If wxParametro555 = "S" Then
                           btnImprimeAtencion_Click
                        End If
                        Me.Visible = False
                        LimpiarVariablesDeMemoria
                   End If
               Else
                   ms_NombrePaciente = ""
                   MsgBox "No se pudo modificar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   'Case sghEliminar
    '   MsgBox "No se puede ELIMINAR desde este módulo", vbInformation, Me.Caption
   End Select
End Sub

Sub ActualizaVacunasDesdeModuloPerinatal()
    If Me.ucPerinatal1.Visible = False Then
       Exit Sub
    End If
    On Error GoTo errPer
    Dim oDOFactOrdenServicio As New DoFactOrdenServ
    Dim oDoCatalogoServicioHosp As New DOFinanciamientoCatalogoServ
    Dim mrs_FacturacionProductos As New Recordset
    Dim oRsVacunas As New Recordset
    Dim oConexion As New Connection
    Dim oRsTmp1 As New Recordset
    Dim oFactOrdenServicio As New FactOrdenServicio
    Dim lnPrecioUnitario As Double, lnIdPuntoCarga As Long
    'mgaray201410e
    Dim oRsProcedimientos As New Recordset
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion
    
    With oDOFactOrdenServicio
         .fechacreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL      'Now
         .idCuentaAtencion = Me.idCuentaAtencion
         .idestadofacturacion = sghEstadoFacturacion.sghAtendido
         .IdFuenteFinanciamiento = ml_IdFuenteFinanciamiento
         .idPaciente = ml_IdPaciente
         .idPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaServicioHospitalizacion   'consumo en el servicio
         .idTipoFinanciamiento = ml_IdFormaPago
         .idUsuario = ml_idUsuario
         .IdUsuarioAuditoria = ml_idUsuario
         .FechaDespacho = .fechacreacion
         .IdUsuarioDespacho = ml_idUsuario
         .FechaHoraRealizaCpt = .fechacreacion
    End With
    'mgaray20141024
    If ml_idOrdenServicioInmunizaciones > 0 Then
        Set oFactOrdenServicio.Conexion = oConexion
        oDOFactOrdenServicio.IdOrden = ml_idOrdenServicioInmunizaciones
        If oFactOrdenServicio.SeleccionarPorId(oDOFactOrdenServicio) = True Then
        End If
        Set oFactOrdenServicio.Conexion = Nothing
    End If
    Set oRsVacunas = Me.ucPerinatal1.DevuelveCptInmunizaciones
    'mgaray201410f
'    Set oRsProcedimientos = Me.ucPerinatal1.DevuelveCptFrecuentes
    If oRsVacunas.RecordCount > 0 Then
        'mgaray201410f
        Set mrs_FacturacionProductos = retornaRsProductoParaCpt()
        'mgaray201410e
        If oRsVacunas.RecordCount > 0 Then
            oRsVacunas.MoveFirst
            Do While Not oRsVacunas.EOF
                 Set oDoCatalogoServicioHosp = mo_AdminFacturacion.CatalogoServiciosHospSeleccionarPorId(oRsVacunas!ID, _
                                                                                                         ml_IdFormaPago, _
                                                                                                         oConexion)
                 lnPrecioUnitario = oDoCatalogoServicioHosp.PrecioUnitario
                 mrs_FacturacionProductos.AddNew
                 mrs_FacturacionProductos.Fields!Codigo = ""
                 mrs_FacturacionProductos.Fields!idProducto = oRsVacunas!ID
                 mrs_FacturacionProductos.Fields!NombreProducto = oRsVacunas!procedimiento
                 mrs_FacturacionProductos.Fields!PrecioUnitario = lnPrecioUnitario
                 mrs_FacturacionProductos.Fields!TotalPorPagar = lnPrecioUnitario
                 mrs_FacturacionProductos.Fields!Cantidad = 1
                 mrs_FacturacionProductos.Fields!idestadofacturacion = 1
                 mrs_FacturacionProductos.Update
                 oRsVacunas.MoveNext
            Loop
        End If
'        If oRsServiciosIntermedios.RecordCount = 0 Then
        If ml_idOrdenServicioInmunizaciones = 0 Or mi_Opcion = sghAgregar Then
            If mo_AdminFacturacion.FactOrdenServicioAgregar(oDOFactOrdenServicio, mrs_FacturacionProductos, _
                                                         mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Me.Caption, _
                                                         ml_IdServicio, 0, 0) = True Then
                'mgaray201410e actualizar la orden en los procedimientos de perinatal
                Call mo_AdminAdmision.PerinatalModicarOrdenServicio(mo_Atenciones.idAtencion, oDOFactOrdenServicio.IdOrden, _
                            ml_idUsuario, oConexion)
            End If
        Else
            If mo_AdminFacturacion.FactOrdenServicioModificar(oDOFactOrdenServicio, mrs_FacturacionProductos, _
                                                         mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Me.Caption) = True Then
                Call mo_AdminAdmision.PerinatalModicarOrdenServicio(mo_Atenciones.idAtencion, oDOFactOrdenServicio.IdOrden, _
                            ml_idUsuario, oConexion)
            End If
        End If
'    ElseIf oRsServiciosIntermedios.RecordCount > 0 Then
    ElseIf oRsServiciosIntermedios.RecordCount > 0 And ml_idOrdenServicioInmunizaciones > 0 And mi_Opcion <> sghAgregar Then
        If mo_AdminFacturacion.FactOrdenServicioEliminar(oDOFactOrdenServicio, mo_lnIdTablaLISTBARITEMS, _
                                                                      mo_lcNombrePc, Me.Caption, 0, 0) = True Then
        End If
    End If
    oConexion.Close
    Set oDOFactOrdenServicio = Nothing
    Set oDoCatalogoServicioHosp = Nothing
    Set mrs_FacturacionProductos = Nothing
    Set oRsVacunas = Nothing
    Set oRsTmp1 = Nothing
    Set oConexion = Nothing
    Set oFactOrdenServicio = Nothing
errPer:
End Sub





Private Sub btnCancelar_Click()
   Dim lbSale As Boolean
   lbSale = False
   If sighEntidades.ParaAuditoria = "" Then
      lbSale = True
   ElseIf MsgBox("Hubo cambios, desea salir de todas maneras ?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
      lbSale = True
   End If
   If lbSale = True Then
        Me.Visible = False
        LimpiarVariablesDeMemoria
        If ml_AScorrelativo > 0 And lb_YaSeRegistroDatos = True Then
            mo_Atenciones.IdUsuarioAuditoria = ml_idUsuario
            mo_CuentasAtencion.IdUsuarioAuditoria = ml_idUsuario
            Set mo_Diagnosticos = Nothing
            Me.UcDiagnosticoDetalle1.CargarDiagnosticosAlObjetoDatosMenosCtaActual mo_Diagnosticos
            If mo_AdminAdmision.AtencionesEnOtrosConsultoriosAlMismoTiempo(mo_paciente, mo_Atenciones, _
                                                                           mo_DoAtencionDatosAdicionales, mo_CuentasAtencion, _
                                                                           mo_DOAtencionesCE, mo_lnIdTablaLISTBARITEMS, _
                                                                           mo_lcNombrePc, mi_Opcion, _
                                                                           Me.UcDiagnosticoDetalle1.DevuelveDx, _
                                                                           Me.UcDiagnosticoDetalle1.TipoDiagnostico, oRsGrdOtrosCpt, True) = False Then
            End If
        End If
    End If
End Sub

Function ValidarDatosObligatorios() As Boolean
   If mi_Opcion = sghEliminar Then
       If wxParametro514 <> "S" Then
          MsgBox "No se podrá ELIMINAR DATOS DE LA ATENCION verifique PARAMETROS", vbInformation, Me.Caption
          ValidarDatosObligatorios = False
          Exit Function
       ElseIf lbYaHuboDespacho = True Then
          MsgBox "No se podrá ELIMINAR DATOS DE LA ATENCION porque hubo despachos", vbInformation, Me.Caption
          ValidarDatosObligatorios = False
          Exit Function
       ElseIf MsgBox("No se olvide de eliminar los CPTS realizados en el CONSULTORIO (botón X en la FICHA 3.2.1)" & Chr(13) & Chr(13) & "está seguro de ELIMINAR LA ATENCION ?", vbQuestion + vbYesNo, "") = vbNo Then
          ValidarDatosObligatorios = False
          Exit Function
       End If
       LimpiaTodosControles
       mo_Atenciones.HoraEgreso = "99:99"
       Me.UcRecetas1.LimpiarDatos False
       ValidarDatosObligatorios = True
       Exit Function
   End If

   Dim sMensaje As String
   ValidarDatosObligatorios = False
         
    '---------------------------------------------------------------------------------
    '           VALIDA DATOS DE LA ATENCION
    '---------------------------------------------------------------------------------
   
    '---------------------------------------------------------------------------------
    '           VALIDA DATOS DE PACIENTES
    '---------------------------------------------------------------------------------
    '
    
'AGREGADO POR FRANK - MODULO PROGRAMA MATERNO
    'VALIDA DATOS DE CONTROL DE PROGRAMA MATERNO
    ms_MensajeError = ""
    If Me.UcProgramaMaterno.Visible = True Then
        If UcProgramaMaterno.ValidarDatosObligatorios = False Then
            ms_MensajeError = UcProgramaMaterno.MensajeValidacion + " (Programa Materno)" + Chr(13)
        End If
   
        If ms_MensajeError <> "" Then
           sMensaje = sMensaje + ms_MensajeError
        End If
    End If
    'si existe el cpt "ad040" debe ingresarse el la columna DETALLE ADICIONAL
    If oRsGrdOtrosCpt.RecordCount > 0 Then
       oRsGrdOtrosCpt.MoveFirst
       Do While Not oRsGrdOtrosCpt.EOF
            If UCase(Trim(oRsGrdOtrosCpt!Codigo)) = lcAD040 Then
               If IsNull(oRsGrdOtrosCpt!labConfHIS) Or oRsGrdOtrosCpt!labConfHIS = "" Then
                  sMensaje = sMensaje & "Existe el CPT " & lcAD040 & " tiene que registrar valor en la columna DETALLE ADICIONAL" & Chr(13)
               End If
            End If
            oRsGrdOtrosCpt.MoveNext
       Loop
    End If
    '
    If Me.ucPerinatalAS1.Visible = True Then
      If Me.ucPerinatalAS1.ValidarDatosObligatorios() = False Then
         Exit Function
      End If
    End If
    '
    If sMensaje <> "" Then
        MsgBox sMensaje, vbInformation, Me.Caption
        Exit Function
    End If
    ValidarDatosObligatorios = True

End Function

Function ValidarReglas() As Boolean
    If mi_Opcion = sghEliminar Then
       ValidarReglas = True
       Exit Function
    End If
    
    If Me.UcRecetas1.ValidaReglas = False Then
       Exit Function
    End If
    
    Dim rsCitas  As Recordset
    Dim lcMensaje As String
    Dim lcFichaEpisodio As String
    ValidarReglas = False
    Dim oDOAtencionesCE As DOAtencionesCE
    Set oDOAtencionesCE = ucTriajeVisorCE.DOAtencionCE
   
    If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE Then
        If Me.ucPerinatal1.Visible = True Then
           If Me.ucPerinatal1.DevuelveDxMorbilidad.RecordCount = 0 Then
                MsgBox "Debe ingresar al menos un Dx de Morbilidad (Modulo Niño Sano)", vbExclamation, Me.Caption
                Me.TabAtencion.Tab = 1
                Me.TabDx.Tab = 1
                Me.ucPerinatal1.EnfoqueTabProcedimientos
                Exit Function
           End If
        End If
        'MODULO PROGRAMA MATERNO
        If Me.UcProgramaMaterno.Visible = True Then
             If UcProgramaMaterno.DevuelveProDiagnosticos.RecordCount = 0 Then
                 MsgBox "Debe ingresar al menos un Dx (MODULO MATERNO)", vbExclamation, Me.Caption
                 Me.TabAtencion.Tab = 1
                 Me.TabDx.Tab = 1
                 Exit Function
             End If
         End If
         If mo_Diagnosticos.Count = 0 Then
                 If Me.ucPerinatal1.Visible = False And Me.UcProgramaMaterno.Visible = False And Me.ucPerinatalAS1.Visible = False Then
                    MsgBox "Debe ingresar al menos un Dx (ficha 3.2)", vbExclamation, Me.Caption
                    Me.TabAtencion.Tab = 1
                    'mgaray201410f
                    If Me.TabDx.TabVisible(0) = True Then
                        Me.TabDx.Tab = 0
                    End If
                    Exit Function
                 End If
           End If
         'debb-09/06/2016 (fin)
           If mb_NecesitaTriaje = True And oDOAtencionesCE Is Nothing Then
                'mgaray20141023
                MsgBox "Debe registrar datos de TRIAJE (temperatura, peso, talla) " & Chr(13) & "El Consultorio está configurado para que se registre TRIAJE" & Chr(13) & "(ficha Atencion - F12)", vbExclamation, Me.Caption
                Exit Function
           End If
           If Not (oDOAtencionesCE Is Nothing) Then
                'mgaray20141022
                If mb_NecesitaTriaje = True And (Trim(oDOAtencionesCE.TriajeTemperatura) = "" And Trim(oDOAtencionesCE.triajePeso) = "" And Trim(oDOAtencionesCE.triajeTalla) = "") Then
                    MsgBox "Debe registrar datos de TRIAJE (temperatura, peso, talla) " & Chr(13) & "El Consultorio está configurado para que se registre TRIAJE" & Chr(13) & "(ficha Atencion - F12)", vbExclamation, Me.Caption
                    Exit Function
                End If
           End If
           If cmbIdDestinoAtencion.Text = "" Then
                MsgBox "Debe elegir DESTINO (Ficha 3.5)", vbExclamation, Me.Caption
                Me.TabAtencion.Tab = 4
                Exit Function
           End If
'           If mb_NecesitaTriaje = True And Trim(txtTemperatura.Text) = "" Then
'                MsgBox "Debe registrar datos de TRIAJE (Presión, temperatura, peso, talla) " & Chr(13) & "El Consultorio está configurado para que se registre TRIAJE" & Chr(13) & "(ficha Atencion - F12)", vbExclamation, Me.Caption
'                Exit Function
'           End If
'           If mb_NecesitaTriaje = True And Trim(txtPeso.Text) = "" Then
'                MsgBox "Debe registrar datos de TRIAJE (Presión, temperatura, peso, talla) " & Chr(13) & "El Consultorio está configurado para que se registre TRIAJE" & Chr(13) & "(ficha Atencion - F12)", vbExclamation, Me.Caption
'                Exit Function
'           End If
'           If mb_NecesitaTriaje = True And Trim(txtTalla.Text) = "" Then
'                MsgBox "Debe registrar datos de TRIAJE (Presión, temperatura, peso, talla) " & Chr(13) & "El Consultorio está configurado para que se registre TRIAJE" & Chr(13) & "(ficha Atencion - F12)", vbExclamation, Me.Caption
'                Exit Function
'           End If
'           If Me.txtPresion.Text <> "___/___" Then
'                Dim lcSistolica As String, lcDiastolica As String
'                lcSistolica = Left(Me.txtPresion.Text, InStr(Me.txtPresion.Text, "/") - 1)
'                lcDiastolica = Mid(Me.txtPresion.Text, InStr(Me.txtPresion.Text, "/") + 1, 100)
'                If Val(lcSistolica) < Val(lcDiastolica) Then
'                    MsgBox "En la Presión: Sistolica debe ser mayor a la Diastólica", vbInformation, Me.Caption
'                    Exit Function
'                End If
'           End If
'           If Me.txtTemperatura.Text <> "" Then
'                If Not (Val(Me.txtTemperatura.Text) >= 35 And Val(Me.txtTemperatura.Text) <= 42) Then
'                    MsgBox "La Temperatura debe estar entre 35 y 42 °C ", vbInformation, Me.Caption
'                    Exit Function
'                End If
'           End If
'
'           If Val(Me.txtPulso.Text) > 250 Then
'                MsgBox "El PULSO no debe pasar de 250" & Chr(13) & "(ficha Atencion - F12)", vbExclamation, Me.Caption
'                Exit Function
'           End If
'           If Val(Me.txtFrespiratoria.Text) > 70 Then
'                MsgBox "La FRECUENCIA RESPIRATORIA no debe pasar de 70" & Chr(13) & "(ficha Atencion - F12)", vbExclamation, Me.Caption
'                Exit Function
'           End If
            If wxParametro302 = "S" And ml_IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                Dim doServicio As New doServicio
                Set doServicio = RetornaServicio(ml_IdServicio)
                If IsNull(doServicio.codigoServicioFUA) Or doServicio.codigoServicioFUA = "" Then
                    MsgBox "Debe configurar el UPS del Servicio para el FUA. " & Chr(13) & "Solicite al Administrador que configure el UPS para el FUA en el módulo GENERAL->SERVICIOS.", vbExclamation, Me.Caption
                    Set doServicio = Nothing
                    Exit Function
                End If
            End If
            If (Val(mo_cmbIdDestinoAtencion.BoundText) = 12 Or Val(mo_cmbIdDestinoAtencion.BoundText) = 13) And Val(Me.txtIdEstablecimientoDestino.Text) = 0 Then
                 If lcBuscaParametro.SeleccionaFilaParametro(362) = "S" Then
                    MsgBox "Debe elegir el ESTABLECIMIENTO destino del Paciente (ficha 3.5)", vbExclamation, Me.Caption
                 Else
                    MsgBox "Debe elegir el ESTABLECIMIENTO destino del Paciente (ficha 3.4)", vbExclamation, Me.Caption
                 End If
                 Me.TabAtencion.Tab = 4 '4 Frank Desactivo de Tratamiento
                 Exit Function
            End If
            If (Val(mo_cmbIdDestinoAtencion.BoundText) = 12 Or Val(mo_cmbIdDestinoAtencion.BoundText) = 13) Then
                If Trim(Me.txtNroReferenciaDestino.Text) = "" Then
                        If lcBuscaParametro.SeleccionaFilaParametro(362) = "S" Then
                           MsgBox "Debe ingresar el N° DE REFERENCIA destino del Paciente (ficha 3.5)", vbExclamation, Me.Caption
                        Else
                           MsgBox "Debe ingresar el N° DE REFERENCIA destino del Paciente (ficha 3.4)", vbExclamation, Me.Caption
                        End If
                        Me.TabAtencion.Tab = 4
                        Exit Function
                End If
                If mo_AdminServiciosComunes.BuscarSiExisteNUMEROREFERENCIA(txtNroReferenciaDestino.Text, ml_idAtencion, _
                                                  IIf(Val(mo_cmbIdDestinoAtencion.BoundText) = 13, True, False)) = True Then
                   Me.TabAtencion.Tab = 4
                   txtNroReferenciaDestino.Text = mo_AdminServiciosComunes.CalculaNUMEROREFERENCIA(IIf(Val(mo_cmbIdDestinoAtencion.BoundText) = 13, True, False))
                   Exit Function
                End If
                If cmbServicioReferenciaD.Text = "" Then
                   MsgBox "Debe elegir el SERVICIO DE LA REFERENCIA", vbInformation, Me.Caption
                   Me.TabAtencion.Tab = 4
                   Exit Function
                End If
                If IsDate(Me.txtFextension.Text) Then
                   If CDate(Me.txtFextension.Text) < ml_ldFechaIngreso Then
                         MsgBox "La fecha de EXTENSION DE LA REFERENCIA no puede ser menor a la  FECHA DE INGRESO ", vbInformation, Me.Caption
                         Me.TabAtencion.Tab = 4
                         Exit Function
                   End If
                End If
                If IsDate(Me.txtFtramite.Text) Then
                   If CDate(Me.txtFtramite.Text) < ml_ldFechaIngreso Then
                         MsgBox "La fecha de TRAMITE DE LA REFERENCIA no puede ser menor a la FECHA DE INGRESO ", vbInformation, Me.Caption
                         Me.TabAtencion.Tab = 4
                         Exit Function
                   End If
                End If
            
            End If
            If mo_ReglasSISgalenhos.SisFUAyaFueEnviadoAlSisLIMA(ml_idCuentaAtencion, ml_IdFormaPago, wxParametro302) = True Then
                Exit Function
            End If
    End If
    If wxParametro302 = "S" And ml_IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS And _
            mi_Opcion = sghEliminar Then
            Set rsCitas = mo_ReglasSISgalenhos.SisFuaAtencionSeleccionarPorCuenta(ml_idCuentaAtencion)
            If rsCitas.RecordCount > 0 Then
               MsgBox "El formato FUA ya fué generado: " & rsCitas.Fields!fuaDisa & "-" & rsCitas!fuaLote & "-" & _
                      rsCitas!FuaNumero & Chr(13) & "Debe eliminar el formato FUA (módulo: SIS, opción: Formato FUA)", _
                      vbInformation, Me.Caption
               Exit Function
            End If
    End If
    If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE And mi_Opcion = sghModificar Then
       lcFichaEpisodio = "(Ficha 3.4)"
       If lcBuscaParametro.SeleccionaFilaParametro(362) = "S" Then lcFichaEpisodio = "(Ficha 3.5)"
       If Me.UcEpisodioClinico1.ValidaReglas(lcFichaEpisodio) = False Then
          Me.TabAtencion.Tab = 4 '4 Frank Desactivo de Tratamiento
          Exit Function
       End If
    End If
    'mgaray201410c
    If Me.ucPerinatal1.Visible = True Then
        Dim sCodigoCie As String
        sCodigoCie = mo_AdminAdmision.ObtenerCodigoCIEParaAtencionCRED()
        'mgaray201411h
'        If mo_AdminAdmision.ValidarIngresoDiagnosticoAtencionCREDFromRs(Me.ucDiagnosticoDetalle1.rsDiagnosticos, _
'                            "CodigoCIE2004", Me.ucPerinatal1.NumeroSesionDesarrollo, sCodigoCie) = False Then
'            Me.TabAtencion.Tab = 1
'            'mgaray201410f
'            If Me.TabDx.TabVisible(0) = True Then
'                Me.TabDx.Tab = 0
'            End If
'            MsgBox "No se puede especificar Diagnostico " & sCodigoCie & ", porque no se ha ejecutado ninguna sesión de desarrollo (ficha 3.2.1)", vbExclamation, Me.Caption
'            Exit Function
'        End If
'        If mo_AdminAdmision.ValidarIngresoDiagnosticoAtencionCREDFromRs(Me.ucPerinatal1.DevuelveDxDesarrollo, _
'                            "CodigoCIE2004", Me.ucPerinatal1.NumeroSesionDesarrollo, sCodigoCie) = False Then
'            Me.TabAtencion.Tab = 1
'            Me.TabDx.Tab = 1
'            MsgBox "No se puede especificar Diagnostico de desarrollo " & sCodigoCie & ", porque no se ha ejecutado ninguna sesión de desarrollo (ficha 3.2.2)", vbExclamation, Me.Caption
'            Exit Function
'        End If
            
        If Me.ucPerinatal1.NumeroSesionDesarrollo <> "" Then
            If mo_AdminAdmision.ValidarLabDiagnosticoAtencionCREDFromRs(Me.ucPerinatal1.DevuelveDxMorbilidad, _
                                "CodigoCIE2004", Me.ucPerinatal1.NumeroSesionDesarrollo, "labConfHIS", sCodigoCie) = False Then
                Me.TabAtencion.Tab = 1
                'mgaray2014112a
'                If Me.TabDx.TabVisible(0) = True Then
'                    Me.TabDx.Tab = 0
'                End If
                'mgaray201412a
                Me.TabAtencion.Tab = 1
                Me.TabDx.Tab = 1
                MsgBox "Lab Ingresado en Diagnostico " & sCodigoCie & ", no corresponde con el control del niño: sesión " & Me.ucPerinatal1.NumeroSesionDesarrollo & " (ficha 3.2.1)", vbExclamation, Me.Caption
                Exit Function
            End If
            If mo_AdminAdmision.ValidarLabDiagnosticoAtencionCREDFromRs(Me.ucPerinatal1.DevuelveDxDesarrollo, _
                                "CodigoCIE2004", Me.ucPerinatal1.NumeroSesionDesarrollo, "labConfHIS", sCodigoCie) = False Then
                Me.TabAtencion.Tab = 1
                Me.TabDx.Tab = 1
                MsgBox "Lab Ingresado en Diagnostico de desarrollo " & sCodigoCie & ", no corresponde con el control del niño: sesión " & Me.ucPerinatal1.NumeroSesionDesarrollo & " (ficha 3.2.2)", vbExclamation, Me.Caption
                Exit Function
            End If
        End If
    End If
    'debb-03/09/2015
    If mo_ReglasFarmacia.RecetaChequeaSiFechaVigenciaEsCorrecta(Me.UcRecetas1.DevuelveFarmacia) = False And mi_Opcion = sghModificar Then
       Exit Function
    End If
    '
    lcMensaje = mo_ReglasSISgalenhos.ReglasDeConsistenciaSISsoloFarmaciaXmonto(mo_Atenciones.idCuentaAtencion, _
                             mo_Atenciones.IdFuenteFinanciamiento, Format(mo_Atenciones.FechaIngreso, "dd/mm/yyyy"), _
                             IIf(lnRecetaFarmacia = 0, sghAgregar, sghModificar), Me.UcRecetas1.DevuelveFarmacia, True)
    If lcMensaje <> "" Then
       MsgBox lcMensaje, vbInformation, ""
       Exit Function
    End If
    '
    ValidarReglas = True
    Set rsCitas = Nothing
End Function

Function RetornaServicio(ml_IdServ As Long) As doServicio
    Dim oConexion As New Connection
    Dim oDoServicio As New doServicio
    Dim oServicio As New Servicios
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion
    
    Set oServicio.Conexion = oConexion
    oDoServicio.IdServicio = ml_IdServ
    If oServicio.SeleccionarPorId(oDoServicio) Then
    End If
    
    oConexion.Close
    Set oConexion = Nothing
    Set oServicio = Nothing
    Set RetornaServicio = oDoServicio
End Function

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
                .IdUsuarioAuditoria = ml_idUsuario
                .idEstado = sghEstadoCuenta.sghAbierto
   End With
   
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LA ATENCION
    '---------------------------------------------------------------------------------
   With mo_Atenciones
            .IdDestinoAtencion = Val(mo_cmbIdDestinoAtencion.BoundText)
            .IdMedicoEgreso = 0
            If lb_YaSeRegistroDatos = False Then
                .HoraEgreso = ""
                .fechaEgreso = 0
            End If
            .IdTipoGravedad = 0
            .IdUsuarioAuditoria = ml_idUsuario
            .IdEstadoAtencion = sghEstadoTabla.sghRegistrado
            .IdTipoCondicionALEstab = Val(mo_cmbIdCondicionEnElEstablecimiento.BoundText)
            .IdTipoCondicionAlServicio = Val(mo_cmbIdCondicionEnElServicio.BoundText)
            .HoraInicioAtencion = lc_HoraQueCargaFormulario
   End With
   With mo_DoAtencionDatosAdicionales
        .IdTipoReferenciaDestino = Val(mo_cmbIdTipoReferenciaDestino.BoundText)
        If .IdTipoReferenciaDestino = 1 Then
             .idEstablecimientoDestino = Val(Me.txtIdEstablecimientoDestino.Tag)
             .IdEstablecimientoNoMinsaDestino = 0
        Else
             .idEstablecimientoDestino = 0
             .IdEstablecimientoNoMinsaDestino = Val(Me.txtIdEstablecimientoDestino.Tag)
        End If
        .NroReferenciaDestino = Trim(txtNroReferenciaDestino.Text)
        .NumeroDeHijos = Val(Me.txtNroHijos.Text)
        .ProximaCita = IIf(lblProximaCita.Caption = "", 0, lblProximaCita.Caption)
        
        'debb-21/08/2015 (inicio)
        If .idSiaSis = 0 Or .SisCodigo = "" And (wxParametro302 = "S" And _
                         mo_Atenciones.IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS) Then
                mo_ReglasSISgalenhos.SisFiliacionesDevuelveKEY lnAfiliacionSIS4, lcSIScodigo, _
                                     mo_paciente.ApellidoPaterno, mo_paciente.ApellidoMaterno, _
                                     mo_paciente.PrimerNombre, mo_paciente.FechaNacimiento, _
                                     lcCodigoEstablecimientoAdscripcionSIS
                .idSiaSis = lnAfiliacionSIS4
                .SisCodigo = lcSIScodigo
                
        End If
        If mo_Atenciones.IdFuenteFinanciamiento <> sghFuenteFinanciamiento.sghFFSIS Then
         .sisAfiliacion = ""
        End If
        'debb-21/08/2015 (fin)
        'debb-21/06/2016 (inicio)
        '.referenciaOservicio
        '.referenciaOidDiagnostico
        .referenciaDservicio = PVcomboBoxDevuelveEleccion(cmbServicioReferenciaD)
        .referenciaDfextension = IIf(txtFextension.Text = sighEntidades.FECHA_VACIA_DMY, 0, txtFextension.Text)
        .referenciaDftramite = IIf(txtFtramite.Text = sighEntidades.FECHA_VACIA_DMY, 0, txtFtramite.Text)
        'debb-21/06/2016 (fin)
        
   End With

    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LA CITA
    '---------------------------------------------------------------------------------
   With mo_Cita
   End With


    
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE DIAGNOSTICOS DE INGRESO
    '---------------------------------------------------------------------------------
    If Me.ucPerinatalAS1.Visible = False And Me.UcProgramaMaterno.Visible = False Then
        Me.UcDiagnosticoDetalle1.idUsuario = ml_idUsuario
        Me.UcDiagnosticoDetalle1.TipoDiagnostico = sghAtencionConsultaExterna
        Me.UcDiagnosticoDetalle1.CargarDiagnosticosAlObjetoDatos mo_Diagnosticos
    End If
    '

    '
End Sub


Sub CargaDatosAtencionJamo()
    Dim oDOAtencionesCETriaje As DOAtencionesCE
    
    With mo_DOAtencionesCE
        .CitaDiagMed = Left(mo_AdminFacturacion.DevuelveDxAltaMedicaTodosDx(mo_Atenciones.idAtencion, 1, ""), 600) & Chr(13) & Chr(10) & lcLineaChar & txtCitaDxMedico.Text
        .CitaDniMedicoJamo = ""
        .CitaExamenClinico = Me.txtCitaExamenClinico.Text
        .CitaExClinicos = "" 'Usado por compatibilidad con Hospital Jmo(Datos Importados de un sistema anterior)
        .CitaFecha = CDate(ml_ldFechaIngreso & " " & ml_lcHoraIngreso)
        .CitaFechaAtencion = lcBuscaParametro.RetornaFechaHoraServidorSQL
        .CitaIdServicio = mo_Atenciones.IdServicioIngreso
        .CitaIdUsuario = mo_Atenciones.IdUsuarioAuditoria
        .CitaMedico = ml_lcMedico
        .CitaMotivo = Me.txtCitaMotivo.Text
        .CitaObservaciones = Me.txtCitaObservaciones.Text
        .CitaServicioJamo = ml_lcServicio
        .CitaTratamiento = Me.TxtCitaTratamiento.Text
        .IdUsuarioAuditoria = mo_Atenciones.IdUsuarioAuditoria
        .NroHistoriaClinica = mo_paciente.NroHistoriaClinica
        .TriajeEdad = ml_lnEdadEnDias
        If IsNull(.idAtencion) Then
           .TriajeFecha = lcBuscaParametro.RetornaFechaHoraServidorSQL
           .TriajeIdUsuario = mo_Atenciones.IdUsuarioAuditoria
        End If
        Call mo_AdminServiciosComunes.cargarDatosTriajeAObjetoDatos(mo_DOAtencionesCE, ucTriajeVisorCE.DOAtencionCE)
        
'        .TriajePeso = Me.txtPeso.Text
'        .TriajePresion = Me.txtPresion.Text
'        .TriajeTalla = Me.txtTalla.Text
'        .TriajeTemperatura = Me.txtTemperatura.Text
'        .TriajePulso = Val(Me.txtPulso.Text)
'        .TriajeFrecRespiratoria = Val(Me.txtFrespiratoria.Text)
        .CitaAntecedente = Me.txtCitaAntecedente.Text
    End With
End Sub


'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------
Function ModificarDatos() As Boolean
        Dim oEpisodioClinico As EpisodioClinico
        oEpisodioClinico = EpisodioClinicoDevuelveDatos
        '
        If mi_Opcion = sghEliminar Then
           mo_Atenciones.HoraEgreso = "99:99"
        End If
        '
        ModificarDatos = mo_AdminAdmision.AdmisionCEModificarAM(mo_Atenciones, mo_Diagnosticos, _
                                           mo_DoAtencionDatosAdicionales, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, _
                                           "", oEpisodioClinico, ml_ldFechaIngreso, mo_Cita.IdCita, oRsSoloLabActividades, _
                                           lb_YaSeRegistroDatos)
        ms_MensajeError = mo_AdminAdmision.MensajeError
        If ms_MensajeError = "" Then

            If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE Then
                If GrabaAtencionJamo = True Then
                   GrabaAtencionPerinatal
                   GrabaAtencionPerinatalAS   'debb-09/06/2016
                   GrabaAtencionProgramaMaterno
                End If
                If ml_FechaReceta = 0 Then
                   ml_FechaReceta = lcBuscaParametro.RetornaFechaHoraServidorSQL
                End If
                '
                ModificarDatos = mo_AdminAdmision.RecetaModificar(ml_idCuentaAtencion, mo_Atenciones.IdServicioIngreso, ml_idUsuario, _
                                                 lnRecetaRayosX, lnRecetaEcografiaO, lnRecetaEcografiaG, lnRecetaTomografia, _
                                                 lnRecetaAnatomiaP, lnRecetaPatologiaC, lnRecetaBancoS, lnRecetaFarmacia, _
                                                 Me.UcRecetas1.DevuelveRayosX, Me.UcRecetas1.DevuelveEcografiaO, _
                                                 Me.UcRecetas1.DevuelveEcografiaG, Me.UcRecetas1.DevuelveTomografia, _
                                                 Me.UcRecetas1.DevuelveAnatomia, Me.UcRecetas1.DevuelvePatologia, _
                                                 Me.UcRecetas1.DevuelveBancoSangre, Me.UcRecetas1.DevuelveFarmacia, ml_FechaReceta, _
                                                 mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "Paciente : " & ms_NombrePaciente, Me.idMedico, False, _
                                                 Me.UcRecetas1.DevuelveOtrosCpt, lnRecetaOtrosCpt)
                ms_MensajeError = mo_AdminAdmision.MensajeError
                If mi_Opcion <> sghEliminar Then
                    If lnRecetaRayosX > 0 Or lnRecetaEcografiaO > 0 Or lnRecetaEcografiaG > 0 Or _
                       lnRecetaTomografia > 0 Or lnRecetaAnatomiaP > 0 Or lnRecetaPatologiaC > 0 Or _
                       lnRecetaBancoS > 0 Or lnRecetaFarmacia > 0 Or lnRecetaOtrosCpt > 0 Or lnRecetaOtrosCpt > 0 Then
                       Me.UcRecetas1.Tratamiento = Trim(TxtCitaTratamiento.Text)
                       Me.UcRecetas1.CargaNumeroDeRecetaEimprime lnRecetaRayosX, lnRecetaEcografiaO, lnRecetaEcografiaG, _
                                       lnRecetaTomografia, lnRecetaAnatomiaP, lnRecetaPatologiaC, lnRecetaBancoS, _
                                       lnRecetaFarmacia, True, lnRecetaOtrosCpt
                    End If
                End If
            End If
            '
            If Val(wxParametro208) <> 7686 Then
                'no se considera los PAGANTES, porque se espera que vaya a CAJA, allí si se considera este proceso
                If mo_Atenciones.IdFormaPago > 1 Then
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
            'debb-27/05/2015
            'If ml_AScorrelativo > 0 Then
            Set mo_Diagnosticos = Nothing
            Me.UcDiagnosticoDetalle1.CargarDiagnosticosAlObjetoDatosMenosCtaActual mo_Diagnosticos
            If mo_AdminAdmision.AtencionesEnOtrosConsultoriosAlMismoTiempo(mo_paciente, mo_Atenciones, _
                                                                           mo_DoAtencionDatosAdicionales, mo_CuentasAtencion, _
                                                                           mo_DOAtencionesCE, mo_lnIdTablaLISTBARITEMS, _
                                                                           mo_lcNombrePc, mi_Opcion, _
                                                                           Me.UcDiagnosticoDetalle1.DevuelveDx, _
                                                                           Me.UcDiagnosticoDetalle1.TipoDiagnostico, _
                                                                           oRsGrdOtrosCpt, False, _
                                                                           ml_AScorrelativo) = False Then
            End If
            'End If
            If wxParametro302 = "S" And mo_Atenciones.IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                If mo_AdminAdmision.ActualizaVariosFUASdeVariosConsultorios(mo_Atenciones.idAtencion, _
                                                     Me.UcDiagnosticoDetalle1.DevuelveDx, oRsGrdOtrosCpt, _
                                                     Me.UcRecetas1.DevuelveRayosX, Me.UcRecetas1.DevuelveEcografiaO, _
                                                     Me.UcRecetas1.DevuelveEcografiaG, Me.UcRecetas1.DevuelveTomografia, _
                                                     Me.UcRecetas1.DevuelveAnatomia, Me.UcRecetas1.DevuelvePatologia, _
                                                     Me.UcRecetas1.DevuelveBancoSangre, Me.UcRecetas1.DevuelveFarmacia, _
                                                     ml_AScorrelativo) Then
                End If
            End If
            '
            If ml_ups = "301202" Then
               AgregaModificaHemoglobina
            End If
       End If
End Function

Sub AgregaModificaHemoglobina()
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim lnIdProductoCpt As Long, lnIdOrden As Long
  lnIdProductoCpt = 3588   'hemoglobina
  lnIdOrden = Val("999" & Trim(Str(Me.idAtencion)))     '999+idatencion
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighEntidades.CadenaConexion
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabResultadosEliminarXidProductoIdOrden"
    Set oParameter = .CreateParameter("@IdProductoCpt", adInteger, adParamInput, 0, lnIdProductoCpt): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
    .Execute
  End With
  Set oCommand = Nothing
  Set oParameter = Nothing
  '
  With oCommand
      .CommandType = adCmdStoredProc
      Set .ActiveConnection = oConexion
      .CommandTimeout = 150
      .CommandText = "LabResultadosPorItemsActualizar"
      Set oParameter = .CreateParameter("@IdProductoCpt", adInteger, adParamInput, 0, lnIdProductoCpt): .Parameters.Append oParameter
      Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
      Set oParameter = .CreateParameter("@ordenXresultado", adInteger, adParamInput, 0, 1): .Parameters.Append oParameter
      Set oParameter = .CreateParameter("@ValorNumero", adCurrency, adParamInput, 0, Null): .Parameters.Append oParameter
      Set oParameter = .CreateParameter("@ValorTexto", adVarChar, adParamInput, 500, Me.txtHemoglobina.Text): .Parameters.Append oParameter
      Set oParameter = .CreateParameter("@ValorCombo", adVarChar, adParamInput, 100, Null): .Parameters.Append oParameter
      Set oParameter = .CreateParameter("@ValorCheck", adVarChar, adParamInput, 1, Null): .Parameters.Append oParameter
      Set oParameter = .CreateParameter("@lnIdRealizaAnalisis", adInteger, adParamInput, 0, sighEntidades.Usuario): .Parameters.Append oParameter
      Set oParameter = .CreateParameter("@lnIdUsuario", adInteger, adParamInput, 0, sighEntidades.Usuario): .Parameters.Append oParameter
      Set oParameter = .CreateParameter("@lcFecha", adDBTimeStamp, adParamInput, 0, CDate(mo_Atenciones.FechaIngreso)): .Parameters.Append oParameter
      
      Set oParameter = .CreateParameter("@IdGrupo", adInteger, adParamInput, 0, Null): .Parameters.Append oParameter
      Set oParameter = .CreateParameter("@IdItem", adInteger, adParamInput, 0, Null): .Parameters.Append oParameter
      Set oParameter = .CreateParameter("@ValorReferencial", adVarChar, adParamInput, 100, Null): .Parameters.Append oParameter
      Set oParameter = .CreateParameter("@ValorMetodo", adVarChar, adParamInput, 50, Null): .Parameters.Append oParameter
      
      .Execute
  End With
  Set oCommand = Nothing
  Set oParameter = Nothing
    
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla CuentasAtencion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlosControles()
        Dim oRecetaCabecera As New RecetaCabecera
        Dim oRsCabeceraReceta As New Recordset
        Dim oRsServiciosAS As New Recordset
        Dim oConexion As New Connection
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighEntidades.CadenaConexion
        '
        btnAceptar.Enabled = True
        '
        '
                
        '1ro:   CARGAR DATOS DE LA CITA
        CargarDatosDeCitasALosControles oConexion
       
        '2do:   CARGAR DATOS DE LA ATENCION
        CargarDatosDelaAtencion oConexion
                
        
        
        '5to:   CARGAR DATOS DE LOS DIAGNOSTICOS POR ATENCION
        Me.UcDiagnosticoDetalle1.idAtencion = Me.idAtencion
        Me.UcDiagnosticoDetalle1.TipoDiagnostico = sghAtencionConsultaExterna
        Me.UcDiagnosticoDetalle1.CargarDatosDeDiagnosticos oConexion
               
        
        'Carga datos de Triaje y Atencion CE (debb-jamo)
        mo_DOAtencionesCE.CitaDiagMed = ""
        CargaAtencionCEJamo oConexion
       
        'Verifica la Cta Atencion
        If lbCargaAlaVezCitaPacienteAtencionDA = False Then
           Set mo_CuentasAtencion = mo_AdminFacturacion.CuentasAtencionSeleccionarPorId(Me.idCuentaAtencion, oConexion)
        End If
        If mo_CuentasAtencion.idEstado <> 1 Then btnAceptar.Enabled = False
        'Ya tuvo movimientos(Farmacia/servicios), no podrá cambiar de plan
        lbYaHuboDespacho = False
        ms_MensajeError = mo_AdminAdmision.VerificaSiTieneMovimientoFarmaciaOservicio(mo_Atenciones.idCuentaAtencion, mo_Atenciones.idTipoServicio, oConexion)
        If ms_MensajeError <> "" Then
           MsgBox ms_MensajeError, vbInformation, Me.Caption
           lbYaHuboDespacho = True
        End If
        ms_MensajeError = ""
        '
        mb_NecesitaTriaje = mo_AdminAdmision.ElServicioNecesitaTriaje(mo_Atenciones.IdServicioIngreso, oConexion, lbElConsultorioUsaModuloPerinatal, lbElConsultorioUsaModuloMaterno)
        '
        CargarDatosPerinatal oConexion
        CargarDatosPerinatalAS oConexion                'debb-09/06/2016
        CargarDatosProgramaMaterno oConexion
        '
        If mo_DoAtencionDatosAdicionales.ProximaCita <> 0 Then
            lblProximaCita.Caption = mo_DoAtencionDatosAdicionales.ProximaCita
            Me.txtProximaCita.Text = mo_DoAtencionDatosAdicionales.ProximaCita - mo_Atenciones.FechaIngreso
        End If
        txtNroHijos.Text = mo_DoAtencionDatosAdicionales.NumeroDeHijos
        lnAfiliacionSIS4 = mo_DoAtencionDatosAdicionales.idSiaSis
        lcSIScodigo = mo_DoAtencionDatosAdicionales.SisCodigo
        '
        Set oRecetaCabecera.Conexion = oConexion
        Set oRsCabeceraReceta = oRecetaCabecera.SeleccionarPorIdCuentaAtencion(mo_Atenciones.idCuentaAtencion)
        Me.UcRecetas1.LimpiarDatos False
        Me.UcRecetas1.idTipoFinanciamiento = mo_Atenciones.IdFormaPago
        Me.UcRecetas1.idTipoSexo = mo_paciente.idTipoSexo
        Me.UcRecetas1.DatoCabeceraReceta = "(N° Cuenta=" & Trim(Str(ml_idCuentaAtencion)) & ") /" & Chr(13) & Chr(10) & ml_lcServicio & "/ Paciente=" & lcHistoriaYpaciente
        Me.UcRecetas1.idCuentaAtencion = ml_idCuentaAtencion 'actualizado 22092014
        ml_FechaReceta = 0
        lnRecetaRayosX = 0: lnRecetaEcografiaO = 0: lnRecetaEcografiaG = 0: lnRecetaTomografia = 0
        lnRecetaAnatomiaP = 0: lnRecetaPatologiaC = 0: lnRecetaBancoS = 0: lnRecetaFarmacia = 0
        lnRecetaOtrosCpt = 0:
        If oRsCabeceraReceta.RecordCount > 0 Then
            ml_FechaReceta = oRsCabeceraReceta.Fields!FechaReceta
            Me.UcRecetas1.CargaDatosAcontroles oRsCabeceraReceta, lnRecetaRayosX, lnRecetaEcografiaO, lnRecetaEcografiaG, lnRecetaTomografia, lnRecetaAnatomiaP, lnRecetaPatologiaC, lnRecetaBancoS, lnRecetaFarmacia, lnRecetaOtrosCpt
        End If
        Me.UcRecetas1.Tratamiento = Trim(TxtCitaTratamiento.Text)
        '
        If mi_Opcion <> sghAgregar And mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE Then
           Me.UcEpisodioClinico1.idPaciente = mo_Atenciones.idPaciente
           Me.UcEpisodioClinico1.idAtencion = mo_Atenciones.idAtencion
           Me.UcEpisodioClinico1.Inicializar
           Me.UcEpisodioClinico1.Limpiar
           Me.UcEpisodioClinico1.CargaEpisodiosHistoricos
           Me.UcEpisodioClinico1.CargarDatosAlosControles oConexion
'           Me.TabAtencion = 4
           TabAtencion = 5
        End If
        '******************debb-27/05/2015 (inicio)****************
        Set oRsServiciosAS = mo_AdminAdmision.ServiciosAtenSimultaneaSeleccionarXidServicio(mo_Atenciones.IdServicioIngreso, oConexion)
        If oRsServiciosAS.RecordCount = 0 Then
            ml_AScorrelativo = 0
        End If
        If oRsServiciosAS.RecordCount > 0 Then
            mo_Atenciones.IdUsuarioAuditoria = ml_idUsuario
            mo_CuentasAtencion.IdUsuarioAuditoria = ml_idUsuario
            If mo_AdminAdmision.AtencionesEnOtrosConsultoriosAlMismoTiempo(mo_paciente, mo_Atenciones, _
                                                                           mo_DoAtencionDatosAdicionales, mo_CuentasAtencion, _
                                                                           mo_DOAtencionesCE, mo_lnIdTablaLISTBARITEMS, _
                                                                           mo_lcNombrePc, sghAgregar, _
                                                                           Me.UcDiagnosticoDetalle1.DevuelveDx, _
                                                                           Me.UcDiagnosticoDetalle1.TipoDiagnostico, , False, 1) = False Then
            End If
            Set mo_Atenciones = mo_AdminAdmision.AtencionesSeleccionarPorId(Me.idAtencion, oConexion)
            Set mo_CuentasAtencion = mo_AdminFacturacion.CuentasAtencionSeleccionarPorId(Me.idCuentaAtencion, oConexion)
        End If
        ml_ups = mo_AdminAdmision.BuscaUPSactualDelPaciente(mo_Atenciones.IdServicioIngreso)
        lb_YaSeRegistroDatos = IIf(mo_Atenciones.HoraEgreso <> "", True, False)
        Me.UcDiagnosticoDetalle1.IdServicio = mo_Atenciones.IdServicioIngreso
        Me.UcDiagnosticoDetalle1.UPS = ml_ups
        Me.UcDiagnosticoDetalle1.FechaAtencion = mo_Atenciones.FechaIngreso
        Me.UcDiagnosticoDetalle1.FechaNacimiento = mo_paciente.FechaNacimiento
        If Val(mo_DOAtencionesCE.triajePeso) >= 0 Then
           Me.UcDiagnosticoDetalle1.PesoKg = Val(mo_DOAtencionesCE.triajePeso)
           lnPesoKg = Val(mo_DOAtencionesCE.triajePeso)
        End If
        Me.UcDiagnosticoDetalle1.IdFuenteFinanciamiento = mo_Atenciones.IdFuenteFinanciamiento
        Me.UcDiagnosticoDetalle1.idCuentaAtencion = ml_idCuentaAtencion
        Me.UcDiagnosticoDetalle1.Consultorio = ml_lcServicio
        If oRsServiciosAS.RecordCount > 0 Then
            ml_AScorrelativo = mo_AdminAdmision.ServiciosAtenSimultaneaMovXidatencion(mo_Atenciones.idAtencion, oConexion)
        End If
        Me.UcDiagnosticoDetalle1.AScorrelativo = ml_AScorrelativo
        Me.UcRecetas1.IdFuenteFinanciamiento = mo_Atenciones.IdFuenteFinanciamiento
        Me.UcRecetas1.AScorrelativo = ml_AScorrelativo
        Me.UcRecetas1.Opcion = IIf(lb_YaSeRegistroDatos = True, sghModificar, sghAgregar)     'debb-09/07/2015
        If (wxParametro302 = "S" And ml_IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS) Then
            If ml_AScorrelativo = 0 Then
                Set oRsServiciosAtenSimultaneaFuaXcorrelativo = mo_AdminAdmision.ServiciosAtenSimultaneaFuaXidatencion(ml_AScorrelativo, mo_Atenciones.idAtencion)
            Else
                Set oRsServiciosAtenSimultaneaFuaXcorrelativo = mo_AdminAdmision.ServiciosAtenSimultaneaFuaXcorrelativo(ml_AScorrelativo)
            End If
            Set Me.UcDiagnosticoDetalle1.RsServiciosAtenSimultaneaFuaXcorrelativo = oRsServiciosAtenSimultaneaFuaXcorrelativo
            Set Me.UcRecetas1.RsServiciosAtenSimultaneaFuaXcorrelativo = oRsServiciosAtenSimultaneaFuaXcorrelativo
        Else
            Set oRsServiciosAtenSimultaneaFuaXcorrelativo = mo_AdminAdmision.ServiciosAtenSimultaneaFuaXidatencion(ml_AScorrelativo, -10)
        End If
        GeneraTmpCPT
        CargaCPTrealizadosEnVariosServicios True
        
        '******************debb-27/05/2015 (fin)****************
        lblMedico.Text = ""
        Set oRsCabeceraReceta = mo_ReglasDeProgMedica.MedicosSeleccionarXIdMedico(mo_Atenciones.IdMedicoIngreso)
        If oRsCabeceraReceta.RecordCount > 0 Then
            lblMedico.Text = "Médico: " & Trim(oRsCabeceraReceta.Fields!ApellidoPaterno) & " " & Trim(oRsCabeceraReceta.Fields!ApellidoMaterno) & _
                                 " " & Trim(oRsCabeceraReceta.Fields!Nombres) & " " & _
                                 IIf(IsNull(oRsCabeceraReceta!DNI), "", " (DNI: " & Trim(oRsCabeceraReceta!DNI) & ")") & _
                                 IIf(IsNull(oRsCabeceraReceta!Colegiatura), "", " (Colegiatura: " & Trim(oRsCabeceraReceta!Colegiatura) & ")") & _
                                 IIf(IsNull(oRsCabeceraReceta!rne), "", " (RNE: " & Trim(oRsCabeceraReceta!rne) & ")") & _
                                 IIf(IsNull(oRsCabeceraReceta!descripcion), "", " (Especialidad: " & Trim(oRsCabeceraReceta!descripcion) & ")")
            ml_lcMedico = ml_lcMedico & IIf(IsNull(oRsCabeceraReceta!rne), "", " (RNE: " & Trim(oRsCabeceraReceta!rne) & ")")
            If oRsCabeceraReceta!EsActivo = False Then
               MsgBox "El Médico no está ACTIVO", vbInformation, ""
               Me.Visible = False
            End If
        End If
        oRsCabeceraReceta.Close
        '
        txtHemoglobina.Text = ""
        txtHemoglobina.Visible = False
        lblHemoglobina.Visible = False
        If ml_ups = "301202" Then
            txtHemoglobina.Visible = True
            lblHemoglobina.Visible = True
            Dim oRsTmp987 As New Recordset
            Set oRsTmp987 = mo_ReglasLaboratorio.LabResultadoPorItemsSeleccionarPorPaciente(mo_Atenciones.idPaciente, 1, oConexion, _
                                                                                    Val("999" & Trim(Str(mo_Atenciones.idAtencion))))
            If oRsTmp987.RecordCount > 0 Then
               oRsTmp987.MoveFirst
               oRsTmp987.Find "fecha<='" & mo_Atenciones.FechaIngreso & "'"
               If Not oRsTmp987.EOF Then
                     If Not IsNull(oRsTmp987!ValorTexto) Then
                        txtHemoglobina.Text = oRsTmp987!ValorTexto
                     ElseIf Not IsNull(oRsTmp987!ValorNumero) Then
                        txtHemoglobina.Text = Trim(Str(oRsTmp987!ValorNumero))
                     End If
               End If
            End If
            oRsTmp987.Close
            Set oRsTmp987 = Nothing
        End If
        '        '
        
        
        oConexion.Close
        Set oConexion = Nothing
        Set oRecetaCabecera = Nothing
        Set oRsCabeceraReceta = Nothing
        Set oRsServiciosAS = Nothing
        '
        If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE Then   'Registro de Atención
           btnCpt.Visible = True
        Else
           btnCpt.Visible = False
        End If
        '

        'mgaray201410f
        Call BloqueoOpcionesMorbilidad
        Call ocultarControles
        '
        If Len(Trim(mo_DOAtencionesCE.CitaDiagMed)) > 0 Then
           MsgBox "Ya se registró la Atención", vbInformation, Me.Caption
        End If
        '
        'franklin 2017
        Me.ucCitasLista1.NOCargaDesdeCitas = True
        Me.ucCitasLista1.IdMedicoAtencion = mo_Atenciones.IdMedicoIngreso
        Me.ucCitasLista1.nroHistoriaCitadoXmedico = mo_paciente.NroHistoriaClinica
        Me.ucCitasLista1.idPacienteCitadoXmedico = mo_Atenciones.idPaciente
        
        CreaYllenaTemporalesActividades False
        
        ucPacientesCtasPDF1.Inicializar mo_Atenciones.idPaciente, mo_Atenciones.idCuentaAtencion
        'carga Imagen..........si demora mucho al cargar, cambiar en parametros la ruta
        Dim lcRutaImg As String
        lcRutaImg = wxParametro237 & "\" & Trim(Str(mo_paciente.NroHistoriaClinica)) & ".jpg"
        If sighEntidades.ArchivoExiste(lcRutaImg) Then
           pi_ImagSeleccionada.Picture = LoadPicture(lcRutaImg)
        Else
           pi_ImagSeleccionada.Picture = LoadPicture("")
        End If
        '
        btnImprimeAtencion.Enabled = True
        If lb_YaSeRegistroDatos = False Then
           btnImprimeAtencion.Enabled = False
        End If
        lbHuboCambioEnDato = False
        
        
        If lbTienePermisoParaImprimirAtencion = False Then
           btnImprimeAtencion.Visible = False
        End If
End Sub



Sub CargarDatosPerinatal(oConexion As Connection)
    If Me.ucPerinatal1.Visible = True Then
       TabDx.Tab = 1
       Dim lnEdadEnDias As Integer, lnIdTipoEdad As Integer
       Dim oDOAtencionesCE As DOAtencionesCE
       
       Set oDOAtencionesCE = RetornaObjetoDatosTriaje()
       
       lnEdadEnDias = ml_lnEdadEnDias
       lnIdTipoEdad = ml_lnIdTipoEdad
       'If lbCargaUnaSolaVez = False Then
           Me.ucPerinatal1.idUsuario = ml_idUsuario
           Me.ucPerinatal1.Inicializar
       'End If
       Me.ucPerinatal1.idPaciente = mo_Atenciones.idPaciente
       Me.ucPerinatal1.idAtencion = mo_Atenciones.idAtencion
       'mgaray201411e
       Me.ucPerinatal1.NroHistoriaClinica = mo_paciente.NroHistoriaClinica
       Set Me.ucPerinatal1.DOAtencionesCE = Me.ucTriajeVisorCE.DOAtencionCE
       'mgaray201410e
       Me.ucPerinatal1.idCuentaAtencion = ml_idCuentaAtencion
       Me.ucPerinatal1.IdFormaPago = ml_IdFormaPago
       Me.ucPerinatal1.IdServicioIngreso = mo_Atenciones.IdServicioIngreso
       Me.ucPerinatal1.idTipoSexo = mo_paciente.idTipoSexo
       Me.ucPerinatal1.EdadEnMeses = sighEntidades.DevuelveEdadEnMeses(mo_paciente.FechaNacimiento, mo_Atenciones.FechaIngreso)
       Me.ucPerinatal1.FechaAtencion = mo_Atenciones.FechaIngreso
       Me.ucPerinatal1.FechaNacimiento = mo_paciente.FechaNacimiento
       Me.ucPerinatal1.CargaDatosAcontroles lnEdadEnDias, lnIdTipoEdad, Val(oDOAtencionesCE.triajePeso), _
                                            Val(oDOAtencionesCE.triajeTalla), oConexion
        Me.ucPerinatal1.cargarDatosAtencionIntegral
'       txtTalla.SetFocus
       SendKeys "{tab}"
    End If
End Sub


'AGREGADO POR FRANK - MODULO PROGRAMA MATERNO 30102014
Sub CargarDatosProgramaMaterno(oConexion As Connection)
    If Me.UcProgramaMaterno.Visible = True Then
        Dim oDOAtencionesCE As DOAtencionesCE
       
       Set oDOAtencionesCE = RetornaObjetoDatosTriaje()
       
       TabDx.Tab = 2
       'If lbCargaUnaSolaVez = False Then
           Me.UcProgramaMaterno.IdPrograma = lnProgramaMaterno 'MODULO MATERNO
           Me.UcProgramaMaterno.idUsuario = ml_idUsuario 'ml_idUsuario
           Me.UcProgramaMaterno.Inicializar
       'End If
       Me.UcProgramaMaterno.idPaciente = mo_Atenciones.idPaciente
        'mgaray201410e
       Me.UcProgramaMaterno.idCuentaAtencion = ml_idCuentaAtencion
       Me.UcProgramaMaterno.IdFormaPago = ml_IdFormaPago
       Me.UcProgramaMaterno.IdServicioIngreso = mo_Atenciones.IdServicioIngreso
       Me.UcProgramaMaterno.idAtencion = mo_Atenciones.idAtencion
       Me.UcProgramaMaterno.FechaAtencion = mo_Atenciones.FechaIngreso
       Me.UcProgramaMaterno.CargaDatosAcontroles Val(oDOAtencionesCE.triajePeso), Replace(oDOAtencionesCE.TriajePresion, "_", ""), Val(oDOAtencionesCE.triajeTalla), ml_lnEdadEnDias, oConexion  'Frank 02092014
       
       'Actualizado 27102014
        If mb_ControlNuevoMaterno = True Then
            Me.UcProgramaMaterno.ControlNuevo = True
        Else
            Me.UcProgramaMaterno.ControlNuevo = False
        End If
       SendKeys "{tab}"
    End If
End Sub


Sub CargarDatosDelaAtencion(oConexion As Connection)
Dim oDoMedico As New DOMedico
Dim oDOEmpleado As New dOEmpleado
Dim oDOEspecialidades As New Collection
Dim lcEstadoAtencion As String
        lc_HoraQueCargaFormulario = lcBuscaParametro.RetornaHoraServidorSQL
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
        
        'idatencion = Trim(Str(Me.idatencion))
        
        grdServicios.cidatencion = Trim(Str(Me.idAtencion))
        
     
        
        If Not mo_Atenciones Is Nothing Then
           With mo_Atenciones
                
                If .HoraInicioAtencion <> "" Then
                   lc_HoraQueCargaFormulario = .HoraInicioAtencion
                End If
                
                Me.idMedico = .IdMedicoIngreso
                Me.idCuentaAtencion = .idCuentaAtencion
                'Carga datos de la atención
                mo_cmbIdDestinoAtencion.BoundText = .IdDestinoAtencion
                mo_cmbIdCondicionEnElEstablecimiento.BoundText = .IdTipoCondicionALEstab
                mo_cmbIdCondicionEnElServicio.BoundText = .IdTipoCondicionAlServicio
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
                ml_ldFechaIngreso = .FechaIngreso
                ml_lcHoraIngreso = .HoraIngreso
                ml_lnEdadEnDias = .Edad
                ml_lnIdTipoEdad = .idTipoEdad
                ml_IdServicio = .IdServicioIngreso
                ml_lcServicio = mo_AdminFacturacion.BuscaServicioActualDelPaciente(ml_IdServicio)
                
                ml_lcMedico = mo_ReglasDeProgMedica.MedicosDevuelveNombre(.IdMedicoIngreso, oConexion)
                ml_IdPaciente = .idPaciente
                ml_IdFuenteFinanciamiento = .IdFuenteFinanciamiento
                ml_IdFormaPago = .IdFormaPago
                '
                mb_ExistenDatos = True
           End With
           '
           If lbCargaAlaVezCitaPacienteAtencionDA = False Then
              Set mo_DoAtencionDatosAdicionales = mo_AdminAdmision.AtencionesDatosAdicionalesSeleccionarPorId(Me.idAtencion, oConexion)
           End If
           With mo_DoAtencionDatosAdicionales
                mo_cmbIdTipoReferenciaDestino.BoundText = .IdTipoReferenciaDestino
                CompletarDatosDelEstablecimientoEnElLoad .idEstablecimientoDestino, .IdEstablecimientoNoMinsaDestino, txtIdEstablecimientoDestino, txtNombreDestinoReferencia, .IdTipoReferenciaDestino
                Me.txtNroReferenciaDestino.Text = .NroReferenciaDestino
                Me.txtNroHijos.Text = .NumeroDeHijos
                If .ProximaCita <> 0 Then
                   Me.txtProximaCita.Text = .ProximaCita - mo_Atenciones.FechaIngreso
                End If
                If Me.txtNroReferenciaDestino.Text <> "" Then
                   HabilitarFrameDestino True
                End If
                'debb-21/06/2016 (inicio)
                PVcomboBoxUbicaPosicion .referenciaDservicio, cmbServicioReferenciaD
                txtFextension.Text = IIf(.referenciaDfextension = 0, sighEntidades.FECHA_VACIA_DMY, .referenciaDfextension)
                txtFtramite.Text = IIf(.referenciaDftramite = 0, sighEntidades.FECHA_VACIA_DMY, .referenciaDftramite)
                'debb-21/06/2016 (fin)
           
           End With
           '
           Me.UcDiagnosticoDetalle1.SexoPaciente = mo_paciente.idTipoSexo
           'Frank 01082014
           Select Case ml_lnIdTipoEdad
           Case 1
                Me.UcDiagnosticoDetalle1.EdadPaciente = ml_lnEdadEnDias * 365
           Case 2
                Me.UcDiagnosticoDetalle1.EdadPaciente = ml_lnEdadEnDias * 12
           Case 3
                Me.UcDiagnosticoDetalle1.EdadPaciente = ml_lnEdadEnDias
           End Select
           
           'ESTOS DATOS SE UTILIZARAN MAS ADELANTE PARA ACTUALIZAR LA UBICACION DE PACIENTE
           Dim oPacientesTmp As New SIGHComun.doPaciente
           Set oPacientesTmp = mo_paciente
           Set mo_Pacientes = mo_paciente
           If Not oPacientesTmp Is Nothing Then
                With oPacientesTmp
                
                                 
                     lcHistoriaYpaciente = "(" & _
                                 HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(oPacientesTmp.NroHistoriaClinica)), False) & _
                                 ") " & _
                                 Trim(oPacientesTmp.ApellidoPaterno) & " " & Trim(oPacientesTmp.ApellidoMaterno) & _
                                 " " & Trim(oPacientesTmp.PrimerNombre)
                    Me.Caption = "(HC: " & _
                                  HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(oPacientesTmp.NroHistoriaClinica)), False) & _
                                  " " & _
                                  Trim(oPacientesTmp.ApellidoPaterno) & " " & Trim(oPacientesTmp.ApellidoMaterno) & _
                                  " " & Trim(oPacientesTmp.PrimerNombre) & ")(Estado: " & lcEstadoAtencion & _
                                  ")(Edad: " & getDescripcionEdad(ml_lnEdadEnDias, ml_lnIdTipoEdad, oPacientesTmp.FechaNacimiento, ml_ldFechaIngreso) & _
                                  ")(T.F: " & mo_AdminServiciosComunes.FuentesFinanciamientoDevuelveDescripcion(mo_Atenciones.IdFuenteFinanciamiento) & _
                                  ")(Gs: " & IIf(IsNull(oPacientesTmp.GrupoSanguineo), "", oPacientesTmp.GrupoSanguineo) & _
                                  ")(Frh: " & IIf(IsNull(oPacientesTmp.FactorRh), "", oPacientesTmp.FactorRh) & ")"
                     mo_DoUbicacionPaciente.IdPaisDomicilio = .IdPaisDomicilio
                     mo_DoUbicacionPaciente.IdCentroPobladoDomicilio = .IdCentroPobladoDomicilio
                     
                     mo_DoUbicacionPaciente.IdPaisProcedencia = .IdPaisProcedencia
                     mo_DoUbicacionPaciente.IdCentroPobladoProcedencia = .IdCentroPobladoProcedencia
                     
                     mo_DoUbicacionPaciente.DireccionDomicilio = .DireccionDomicilio
                     
                     txtEdad1.Text = mo_Atenciones.Edad
                     cmbTipoEdad1.ListIndex = mo_Atenciones.idTipoEdad - 1
                     
                End With
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
                      
End Sub



































Private Sub txtNroReferenciaDestino_Change()
lbHuboCambioEnDato = True
End Sub

Private Sub txtNroReferenciaDestino_LostFocus()
    If lbHuboCambioEnDato = True Then
      sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, txtNroReferenciaDestino.Text
      lbHuboCambioEnDato = False
    End If
End Sub

Private Sub txtProximaCita_Change()
lbHuboCambioEnDato = True
End Sub





Private Sub txtProximaCita_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txtProximaCita_LostFocus
    End If
End Sub

'debb-16/05/2016
Private Sub txtProximaCita_LostFocus()
    If lbHuboCambioEnDato = True Then
      sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, txtProximaCita.Text
      lbHuboCambioEnDato = False
    End If
    
     lblProximaCita.Caption = ""
     If Val(txtProximaCita.Text) > 0 Then
        lblProximaCita.Caption = Format(ml_ldFechaIngreso + Val(txtProximaCita.Text), "dd/mm/yyyy")
     ElseIf Val(txtProximaCita.Text) < 0 Then
         txtProximaCita.Text = ""
     End If
End Sub



Private Sub ucDiagnosticoDetalle1_SePresionoTeclaEspecial(KeyCode As Integer)
    If KeyCode = vbKeyF2 Then
       btnAceptar_Click
    End If
End Sub













Sub LimpiarVariablesDeMemoria()

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

        txtCitaExClinicos.Text = Me.UcRecetas1.DevuelveRecetaAntesDeImprimir
        Select Case mi_Opcion
        Case sghAgregar
             GrabaAtencionJamo = True
        Case sghModificar
             CargaDatosAtencionJamo
             Set oRsTmpBuscaAtencion = mo_AdminAdmision.AtencionCESeleccionarPorIdAtencion(ml_idAtencion)
             mo_DOAtencionesCE.idAtencion = ml_idAtencion
             If oRsTmpBuscaAtencion.RecordCount = 0 Then
                GrabaAtencionJamo = mo_AdminAdmision.AtencionCEAgregar(mo_DOAtencionesCE, mo_lnIdTablaLISTBARITEMS, _
                                             mo_lcNombrePc, "IdAtencion: " & Trim(Str(ml_idAtencion)) & "(desde Atención)")
             Else
                GrabaAtencionJamo = mo_AdminAdmision.AtencionCEModificar(mo_DOAtencionesCE, mo_lnIdTablaLISTBARITEMS, _
                                             mo_lcNombrePc, "IdAtencion: " & Trim(Str(ml_idAtencion)) & "(desde Atención)")
             End If
             PacienteDatosAdicionalesGrabar
             'debb-2/3/2015**inicio
             If Val(mo_DOAtencionesCE.triajeTalla) > 0 And Val(mo_DOAtencionesCE.triajePeso) > 0 Then
                Dim oConexion As New Connection
                Dim lnIdDxNutricional As Long, lnGrafXedadEnMeses As Long, lnGrafYpercentilTE As Long, lnGrafYpercentilPT As Long
                Dim lnGrafYpercentilPE As Long, lnZetaPT As Double, lnZetaTE As Double, lnZetaPE As Double
                Dim lnPercentilIMC As Double, lnPercentilIMC_Z As Double
                Dim lnPesoKg As Double, lnTallaCM As Long, ml_EdadEnMeses As Long, lnEdadEnAniosEnAtencion As Integer
                Dim ml_idTipoSexo As Long
                oConexion.CommandTimeout = 300
                oConexion.CursorLocation = adUseClient
                oConexion.Open sighEntidades.CadenaConexion
                lnPesoKg = Val(mo_DOAtencionesCE.triajePeso)
                lnTallaCM = Val(mo_DOAtencionesCE.triajeTalla)
                ml_EdadEnMeses = sighEntidades.DevuelveEdadEnMeses(mo_paciente.FechaNacimiento, mo_Atenciones.FechaIngreso)
                lnEdadEnAniosEnAtencion = IIf(mo_Atenciones.idTipoEdad = 1, mo_Atenciones.Edad, 0)
                ml_idTipoSexo = mo_paciente.idTipoSexo
                lnIdDxNutricional = 0
                lnGrafXedadEnMeses = ml_EdadEnMeses
                lnGrafYpercentilTE = 0
                lnGrafYpercentilPT = 0
                lnGrafYpercentilPE = 0
                lnPercentilIMC = 0
                lnZetaPT = 0
                lnZetaTE = 0
                lnZetaPE = 0
                lnPercentilIMC_Z = 0
                
                Dim oProcesos As New Procesos
                oProcesos.CalculaPercentiles lnPesoKg, lnTallaCM, ml_EdadEnMeses, _
                                             ml_idTipoSexo, lnEdadEnAniosEnAtencion, _
                                             lnGrafYpercentilPE, lnGrafYpercentilTE, lnGrafYpercentilPT, lnPercentilIMC, _
                                             lnZetaPE, lnZetaTE, lnZetaPT, lnPercentilIMC_Z
                Set oProcesos = Nothing
                If lnZetaPE >= -1 And lnZetaPE <= 1 Then
                   lnIdDxNutricional = 1
                ElseIf lnZetaPE >= -2 And lnZetaPE < -1 Then
                   lnIdDxNutricional = 2
                ElseIf lnZetaPE < -2 Then
                   lnIdDxNutricional = 4
                ElseIf lnZetaPE >= 1 And lnZetaPE <= 2 Then
                   lnIdDxNutricional = 1
                ElseIf lnZetaPE > 2 And lnZetaPE <= 3 Then
                   lnIdDxNutricional = 5
                ElseIf lnZetaPE > 3 Then
                   lnIdDxNutricional = 6
                End If
                mo_AdminServiciosComunes.ActualizaTablaPacientesMovimientos oConexion, mo_Atenciones.FechaIngreso, _
                                                   mo_Atenciones.idCuentaAtencion, sghRegistroAtencionCE, False, Val(mo_DOAtencionesCE.triajePeso), _
                                                   Val(mo_DOAtencionesCE.triajeTalla), lnIdDxNutricional, lnGrafXedadEnMeses, _
                                                   lnGrafYpercentilTE, lnGrafYpercentilPT, lnGrafYpercentilPE, lnZetaPT, lnZetaTE, _
                                                   lnZetaPE
                oConexion.Close
                Set oConexion = Nothing
             End If
             'debb-2/3/2015****final
        Case sghEliminar
             CargaDatosAtencionJamo
             GrabaAtencionJamo = mo_AdminAdmision.AtencionCEeliminar(mo_DOAtencionesCE, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "IdAtencion: " & Trim(Str(ml_idAtencion)) & "(desde Atención)")
             
        End Select
        If GrabaAtencionJamo = False Then
           MsgBox "Falló al Grabar DATOS JAMO" & Chr(13) & mo_AdminAdmision.MensajeError
        End If
        Set oRsTmpBuscaAtencion = Nothing
    End If
End Function

Private Function setDatosAtencionIntegral()
    Set mo_AdminAdmision.setRsAtenInteCrecimientoCabecera = Me.ucPerinatal1.getRsCrecimientoPendiente
    Set mo_AdminAdmision.setRsAtenInteDesarrolloCabecera = Me.ucPerinatal1.getRsDesarrolloPendiente
    
    Set mo_AdminAdmision.setRsAtenInteInmunizaciones = Me.ucPerinatal1.getAtencionIntegralInmunizaciones()
    Set mo_AdminAdmision.setRsAtenInteCrecimiento = Me.ucPerinatal1.getAtencionIntegralCrecimiento()
    Set mo_AdminAdmision.setRsAtenInteDesarrollo = Me.ucPerinatal1.getAtencionIntegralDesarrollo()
    Set mo_AdminAdmision.setRsAtenInteSuplemento = Me.ucPerinatal1.getAtencionIntegralSuplemento()
    Set mo_AdminAdmision.setRsAtenInteTamizaje = Me.ucPerinatal1.getAtencionIntegralTamizaje()
    
    Set mo_AdminAdmision.setRsPlanCrecimiento = Me.ucPerinatal1.getPlanCrecimiento
    Set mo_AdminAdmision.setRsPlanDesarrollo = Me.ucPerinatal1.getPlanDesarrollo
    Set mo_AdminAdmision.setRsPlanInmunizaciones = Me.ucPerinatal1.getPlanInmunizaciones
    Set mo_AdminAdmision.setRsPlanSuplemento = Me.ucPerinatal1.getPlanSuplemento
    Set mo_AdminAdmision.setRsPlanTamizaje = Me.ucPerinatal1.getPlanTamizaje
End Function

Function GrabaAtencionPerinatal() As Boolean
    If ucPerinatal1.Visible = True Then
        Dim oDoPerinatalAtencion As New DoPerinatalAtencion
        Select Case mi_Opcion
        Case sghAgregar
             GrabaAtencionPerinatal = True
        Case sghModificar
             Set oDoPerinatalAtencion = Me.ucPerinatal1.DevuelveDatosGenerales
             oDoPerinatalAtencion.idPaciente = mo_Atenciones.idPaciente
             oDoPerinatalAtencion.IdUsuarioAuditoria = mo_Atenciones.IdUsuarioAuditoria
             
             Call setDatosAtencionIntegral
             
             If oDoPerinatalAtencion.idPerinatalAtencion = 0 Then
                GrabaAtencionPerinatal = mo_AdminAdmision.PerinatalCEAgregar(oDoPerinatalAtencion, _
                                                          Me.ucPerinatal1.DevuelveCptInmunizaciones, _
                                                          Me.ucPerinatal1.DevuelveCptFrecuentes, _
                                                          Me.ucPerinatal1.DevuelveDxDesarrollo, _
                                                          Me.ucPerinatal1.DevuelveDxMorbilidad, _
                                                          Me.ucPerinatal1.DevuelveMedicamentos, _
                                                          Me.ucPerinatal1.DevuelveDatosCred, _
                                                          Me.ucPerinatal1.DevuelvePerinatalAtencionCred1, _
                                                          mo_Atenciones.idAtencion)
             Else
                GrabaAtencionPerinatal = mo_AdminAdmision.PerinatalCEModificar(oDoPerinatalAtencion, _
                                                          Me.ucPerinatal1.DevuelveCptInmunizaciones, _
                                                          Me.ucPerinatal1.DevuelveCptFrecuentes, _
                                                          Me.ucPerinatal1.DevuelveDxDesarrollo, _
                                                          Me.ucPerinatal1.DevuelveDxMorbilidad, _
                                                          Me.ucPerinatal1.DevuelveMedicamentos, _
                                                          Me.ucPerinatal1.DevuelveDatosCred, _
                                                          Me.ucPerinatal1.DevuelvePerinatalAtencionCred1, _
                                                          mo_Atenciones.idAtencion)
             End If
        Case sghEliminar
            oDoPerinatalAtencion.idPaciente = mo_Atenciones.idPaciente
            GrabaAtencionPerinatal = mo_AdminAdmision.PerinatalCEeliminar(oDoPerinatalAtencion, mo_Atenciones.idAtencion)
        End Select
        If GrabaAtencionPerinatal = False Then
           MsgBox "Falló al Grabar PERINATAL" & Chr(13) & mo_AdminAdmision.MensajeError
        End If
        Set oDoPerinatalAtencion = Nothing
    End If
End Function

'ACTUALIZADO 28102014
Function GrabaAtencionProgramaMaterno() As Boolean
    If Me.UcProgramaMaterno.Visible = True Then
        Dim oDoProCabecera As New DoProCabecera
        Dim oDOProControles As New DOProControles
        Select Case mi_Opcion
        Case sghAgregar
             GrabaAtencionProgramaMaterno = True
        Case sghModificar
            If UcProgramaMaterno.EsControlActual Then
                Set oDoProCabecera = Me.UcProgramaMaterno.DevuelveProCabecera
                oDoProCabecera.IdUsuarioAuditoria = mo_Atenciones.IdUsuarioAuditoria
                If oDoProCabecera.IdProCabecera = 0 Then
                    GrabaAtencionProgramaMaterno = mo_ReglasComunes.ProgramaControlesAgregar(oDoProCabecera, UcProgramaMaterno.DevuelveDatosProCabecera, UcProgramaMaterno.DevuelveProControles, UcProgramaMaterno.DevuelveDatosProControles, UcProgramaMaterno.DevuelveProDiagnosticos, UcProgramaMaterno.DevuelveProProcedimientos, UcProgramaMaterno.DevuelveProTratamientos, mo_Atenciones.idAtencion, UcProgramaMaterno.DevuelveProHistorialControles)
                Else
                    GrabaAtencionProgramaMaterno = mo_ReglasComunes.ProgramaControlesModificar(oDoProCabecera, UcProgramaMaterno.DevuelveDatosProCabecera, UcProgramaMaterno.DevuelveProControles, UcProgramaMaterno.DevuelveDatosProControles, UcProgramaMaterno.DevuelveProDiagnosticos, UcProgramaMaterno.DevuelveProProcedimientos, UcProgramaMaterno.DevuelveProTratamientos, mo_Atenciones.idAtencion, UcProgramaMaterno.DevuelveEsControlParaActualizar, UcProgramaMaterno.DevuelveProHistorialControles)
                End If
            Else
                GrabaAtencionProgramaMaterno = True
            End If
        Case sghEliminar
            If UcProgramaMaterno.EsControlActual Then
                GrabaAtencionProgramaMaterno = mo_ReglasComunes.ProgramaControlesEliminar(UcProgramaMaterno.DevuelveProControles, 1, UcProgramaMaterno.DevuelveEsControlParaActualizar)
            Else
                GrabaAtencionProgramaMaterno = True
            End If
        End Select
        If GrabaAtencionProgramaMaterno = False Then
           MsgBox "Falló al Grabar Programa Materno" & Chr(13) & mo_AdminAdmision.MensajeError
        End If
        Set oDoProCabecera = Nothing
        Set oDOProControles = Nothing
    End If
End Function

'debb-Jamo
Sub CargaAtencionCEJamo(oConexion As Connection)
       On Error GoTo ErrJamo
       Dim oAtencionesCE As New AtencionesCE
       Dim oConexionExterna As New Connection
       oConexionExterna.CommandTimeout = 300
       oConexionExterna.CursorLocation = adUseClient
       oConexionExterna.Open wxParametroJAMO
       mo_DOAtencionesCE.idAtencion = Me.idAtencion   'ml_idAtencion
       Set oAtencionesCE.Conexion = oConexionExterna
       PacienteDatosAdicionalesCargar oConexion
       
       'mgaray
       ucTriajeVisorCE.Origen = ConsultaExterna
       ucTriajeVisorCE.EstadoPaciente = 0
       ucTriajeVisorCE.OpcionFormulario = mi_Opcion
       ucTriajeVisorCE.AsignarIdAtencionYLlenarControles (mo_Atenciones.idAtencion)
       
       If oAtencionesCE.SeleccionarPorId(mo_DOAtencionesCE) = False Then
           mo_DOAtencionesCE.idAtencion = 0
           Exit Sub
       End If
'       mo_Formulario.HabilitarDeshabilitar Me.txtPeso, True
'       mo_Formulario.HabilitarDeshabilitar Me.txtPresion, True
'       mo_Formulario.HabilitarDeshabilitar Me.txtTalla, True
'       mo_Formulario.HabilitarDeshabilitar Me.txtTemperatura, True
       mo_Formulario.HabilitarDeshabilitar txtCitaMotivo, True
       mo_Formulario.HabilitarDeshabilitar txtCitaExamenClinico, True
       mo_Formulario.HabilitarDeshabilitar TxtCitaTratamiento, True
       mo_Formulario.HabilitarDeshabilitar txtCitaObservaciones, True
'       mo_Formulario.HabilitarDeshabilitar Me.txtFrespiratoria, True
'       mo_Formulario.HabilitarDeshabilitar Me.txtPulso, True
       mo_Formulario.HabilitarDeshabilitar Me.txtCitaAntecedente, True
       mo_Formulario.HabilitarDeshabilitar Me.txtantecedAlergico, True
       mo_Formulario.HabilitarDeshabilitar Me.txtantecedFamiliar, True
       mo_Formulario.HabilitarDeshabilitar Me.txtantecedObstetrico, True
       mo_Formulario.HabilitarDeshabilitar Me.txtantecedPatologico, True
       mo_Formulario.HabilitarDeshabilitar Me.txtantecedQuirurgico, True
       If Not mo_DOAtencionesCE Is Nothing Then
            With mo_DOAtencionesCE
'                 Me.txtPeso.Text = .TriajePeso
'                 Me.txtPresion.Text = .TriajePresion
'                 Me.txtTalla.Text = .TriajeTalla
'                 Me.txtTemperatura.Text = .TriajeTemperatura
                 txtCitaMotivo.Text = .CitaMotivo
                 txtCitaExamenClinico.Text = .CitaExamenClinico
                 TxtCitaTratamiento.Text = .CitaTratamiento
                 txtCitaObservaciones.Text = .CitaObservaciones
'                 Me.txtFrespiratoria.Text = .TriajeFrecRespiratoria
'                 Me.txtPulso.Text = .TriajePulso
                 Me.txtCitaDxMedico.Text = Mid(.CitaDiagMed, InStr(.CitaDiagMed, lcLineaChar) + 1, 1000)
                 Me.txtCitaAntecedente = .CitaAntecedente
            End With
'            If Len(Trim(mo_DOAtencionesCE.CitaDiagMed)) > 0 Then
'               MsgBox "Ya se registró la Atención", vbInformation, Me.Caption
'            End If
            '
            
            '
'            txtTemperatura_LostFocus
       Else
           mo_DOAtencionesCE.idAtencion = 0
       End If
       oConexionExterna.Close
       Set oConexionExterna = Nothing
       txtCitaExClinicos.Text = Me.UcRecetas1.DevuelveRecetaAntesDeImprimir
       Exit Sub
ErrJamo:

'       mo_Formulario.HabilitarDeshabilitar Me.txtPeso, False
'       mo_Formulario.HabilitarDeshabilitar Me.txtPresion, False
'       mo_Formulario.HabilitarDeshabilitar Me.txtTalla, False
'       mo_Formulario.HabilitarDeshabilitar Me.txtTemperatura, False
       mo_Formulario.HabilitarDeshabilitar txtCitaMotivo, False
       mo_Formulario.HabilitarDeshabilitar txtCitaExamenClinico, False
       mo_Formulario.HabilitarDeshabilitar TxtCitaTratamiento, False
       mo_Formulario.HabilitarDeshabilitar txtCitaObservaciones, False
'       mo_Formulario.HabilitarDeshabilitar Me.txtFrespiratoria, False
'       mo_Formulario.HabilitarDeshabilitar Me.txtPulso, False
       mo_Formulario.HabilitarDeshabilitar Me.txtAntecedentes, False
       mo_Formulario.HabilitarDeshabilitar Me.txtantecedAlergico, False
       mo_Formulario.HabilitarDeshabilitar Me.txtantecedFamiliar, False
       mo_Formulario.HabilitarDeshabilitar Me.txtantecedObstetrico, False
       mo_Formulario.HabilitarDeshabilitar Me.txtantecedPatologico, False
       mo_Formulario.HabilitarDeshabilitar Me.txtantecedQuirurgico, False
       mo_Formulario.HabilitarDeshabilitar Me.txtCitaAntecedente, False
       Set oConexionExterna = Nothing
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







Sub PacienteDatosAdicionalesCargar(oConexion As Connection)
    lbPacienteDatosAdicionalesEsNuevo = True
    Set oDoPacienteDatosAdd = mo_AdminAdmision.PacientesDatosAdicionalesSeleccionarPorId(ml_IdPaciente, oConexion)
    If oDoPacienteDatosAdd.idPaciente > 0 Then
       With oDoPacienteDatosAdd
          Me.txtAntecedentes.Text = .antecedentes
          Me.txtantecedAlergico.Text = .antecedAlergico
          Me.txtantecedObstetrico.Text = .antecedObstetrico
          Me.txtantecedQuirurgico.Text = .antecedQuirurgico
          Me.txtantecedFamiliar.Text = .antecedFamiliar
          Me.txtantecedPatologico.Text = .antecedPatologico
       End With
       lbPacienteDatosAdicionalesEsNuevo = False
    End If
End Sub

Sub PacienteDatosAdicionalesGrabar()
    
    Dim oPacientesDatosAdd As New PacientesDatosAdd
    Dim oConexion As New Connection
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oDoPacienteDatosAdd.idPaciente = ml_IdPaciente
    oDoPacienteDatosAdd.antecedentes = Me.txtAntecedentes.Text
    oDoPacienteDatosAdd.antecedAlergico = Me.txtantecedAlergico.Text
    oDoPacienteDatosAdd.antecedObstetrico = Me.txtantecedObstetrico.Text
    oDoPacienteDatosAdd.antecedQuirurgico = Me.txtantecedQuirurgico.Text
    oDoPacienteDatosAdd.antecedFamiliar = Me.txtantecedFamiliar.Text
    oDoPacienteDatosAdd.antecedPatologico = Me.txtantecedPatologico.Text
    oDoPacienteDatosAdd.IdUsuarioAuditoria = ml_idUsuario
    Set oPacientesDatosAdd.Conexion = oConexion
    If lbPacienteDatosAdicionalesEsNuevo = True Then
       If oPacientesDatosAdd.Insertar(oDoPacienteDatosAdd) = False Then
          MsgBox oPacientesDatosAdd.MensajeError, vbInformation, "CE"
       End If
    Else
       If oPacientesDatosAdd.Modificar(oDoPacienteDatosAdd) = False Then
          MsgBox oPacientesDatosAdd.MensajeError, vbInformation, "CE"
       End If
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oPacientesDatosAdd = Nothing
End Sub

Sub InicilizarParametros()
    wxParametro208 = lcBuscaParametro.SeleccionaFilaParametro(208)
    wxParametro211 = lcBuscaParametro.SeleccionaFilaParametro(211)
    wxParametro216 = lcBuscaParametro.SeleccionaFilaParametro(216)
    wxParametro237 = lcBuscaParametro.SeleccionaFilaParametro(237)
    wxParametro258 = lcBuscaParametro.SeleccionaFilaParametro(258)
    wxParametro274 = lcBuscaParametro.SeleccionaFilaParametro(274)
    wxParametro275 = lcBuscaParametro.SeleccionaFilaParametro(275)
    wxParametro276 = lcBuscaParametro.SeleccionaFilaParametro(276)
    wxParametro282 = lcBuscaParametro.SeleccionaFilaParametro(282)
    wxParametro296 = lcBuscaParametro.SeleccionaFilaParametro(296)
    wxParametro302 = lcBuscaParametro.SeleccionaFilaParametro(302)
    wxParametro306 = lcBuscaParametro.SeleccionaFilaParametro(306)
    wxParametro312 = lcBuscaParametro.SeleccionaFilaParametro(312)
    wxParametro322 = lcBuscaParametro.SeleccionaFilaParametro(322)
    wxParametro323 = lcBuscaParametro.SeleccionaFilaParametro(323)
    wxParametro333 = lcBuscaParametro.SeleccionaFilaParametro(333)
    wxParametro237 = lcBuscaParametro.SeleccionaFilaParametro(237)
    wxParametro358 = lcBuscaParametro.SeleccionaFilaParametro(358)
    wxParametro359 = lcBuscaParametro.SeleccionaFilaParametro(359)
    wxParametro362 = lcBuscaParametro.SeleccionaFilaParametro(362)
    wxParametro502 = lcBuscaParametro.SeleccionaFilaParametro(502)            'debb-09/06/2016
    wxParametro514 = lcBuscaParametro.SeleccionaFilaParametro(514)
    wxParametro518 = lcBuscaParametro.SeleccionaFilaParametro(518)   'franklin 2017
    wxParametro542 = lcBuscaParametro.SeleccionaFilaParametro(542)
    wxParametro545 = lcBuscaParametro.SeleccionaFilaParametro(545)
    wxParametro553 = lcBuscaParametro.SeleccionaFilaParametro(553)
    wxParametro555 = lcBuscaParametro.SeleccionaFilaParametro(555)
    wxParametroJAMO = lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    ldFechaActualServidor = lcBuscaParametro.RetornaFechaServidorSQL
    lcAD040 = lcBuscaParametro.SeleccionaFilaParametro(354)
End Sub




Function EpisodioClinicoDevuelveDatos() As EpisodioClinico
        Dim oEpisodioClinico As EpisodioClinico
        oEpisodioClinico.idEpisodio = Me.UcEpisodioClinico1.idEpisodio
        oEpisodioClinico.lbCierreEpisodio = Me.UcEpisodioClinico1.lbCierreEpisodio
        oEpisodioClinico.lbNuevoEpisodio = Me.UcEpisodioClinico1.lbNuevoEpisodio
        EpisodioClinicoDevuelveDatos = oEpisodioClinico
End Function



Sub HabilitarFrameDestino(bValue As Boolean)
        mo_Formulario.HabilitarDeshabilitar fraDatosReferenciaDestino, bValue
        mo_Formulario.HabilitarDeshabilitar fraDatosReferenciaDestino, bValue
        mo_Formulario.HabilitarDeshabilitar lblIdTipoReferenciaDestino, bValue
        mo_Formulario.HabilitarDeshabilitar cmbIdTipoReferenciaDestino, bValue
        mo_Formulario.HabilitarDeshabilitar txtNroReferenciaDestino, bValue
        mo_Formulario.HabilitarDeshabilitar lblIdEstablecimientoDestino, bValue
        'debb-21/06/2016 (inicio)
        mo_Formulario.HabilitarDeshabilitar cmbServicioReferenciaD, bValue
        mo_Formulario.HabilitarDeshabilitar txtFextension, bValue
        mo_Formulario.HabilitarDeshabilitar txtFtramite, bValue
        mo_Formulario.HabilitarDeshabilitar Me.lblFtramite0, bValue
        mo_Formulario.HabilitarDeshabilitar Me.lblFextension0, bValue
        mo_Formulario.HabilitarDeshabilitar lblServicioReferencia0, bValue
        mo_Formulario.HabilitarDeshabilitar lblReferenciaO, bValue
        'debb-21/06/2016 (fin)
End Sub



Private Sub ucPerinatal1_LostFocus()
      Me.UcRecetas1.CargaRecetaDesdePerinatal Me.ucPerinatal1.DevuelveMedicamentos, Me.ucPerinatal1.DevuelveCptFrecuentes
      
End Sub




'debb-09/06/2016
Private Sub ucPerinatalAS1_LostFocus()
    Me.UcRecetas1.CargaRecetaDesdePerinatal Me.ucPerinatalAS1.DevuelveMedicamentos, Me.ucPerinatalAS1.DevuelveCptFrecuentes
    
End Sub

Private Sub UcProgramaMaterno_LostFocus()
      Me.UcRecetas1.CargaRecetaDesdeMaterno UcProgramaMaterno.DevuelveProTratamientos, UcProgramaMaterno.DevuelveProProcedimientos
End Sub



Private Sub ucTriajeVisorCE_changeDataControl(mo_DOAtencionesCEActual As SIGHComun.DOAtencionesCE, mo_DOAtencionesCENew As SIGHComun.DOAtencionesCE)
    '===============================================
    'peso_lostfocus
    'mgaray201411e
    If ucPerinatal1.Visible = True And Val(mo_DOAtencionesCENew.triajePeso) > 0 Then
'       ucPerinatal1.ActualizaGraficoYDiagnosticosAutomaticamente Val(mo_DOAtencionesCENew.TriajePeso), Val(mo_DOAtencionesCENew.TriajeTalla)
       ucPerinatal1.ActualizaGraficoYDiagnosticosAutomaticamente mo_DOAtencionesCENew
    End If
    'debb-09/06/2016 (inicio)
    If ucPerinatalAS1.Visible = True Then
       If Val(mo_DOAtencionesCENew.triajePeso) > 0 Or Val(mo_DOAtencionesCENew.triajeTalla) > 0 Then
          ucPerinatalAS1.ActualizaGraficoYDiagnosticosAutomaticamente Val(mo_DOAtencionesCENew.triajePeso), Val(mo_DOAtencionesCENew.triajeTalla)
          ucPerinatalAS1.LimpiaDxDesarrollo
          ucPerinatalAS1.CargaDxAutomaticosParaMorbilidadEnDesarrollo Val(mo_DOAtencionesCENew.triajePeso), Val(mo_DOAtencionesCENew.triajeTalla)
       End If
    End If
    'debb-09/06/2016 (fin)
    If UcProgramaMaterno.Visible = True And Val(mo_DOAtencionesCENew.triajePeso) > 0 And Val(mo_DOAtencionesCENew.triajeTalla) > 0 Then
       UcProgramaMaterno.Actualiza_Peso_Presion Val(mo_DOAtencionesCENew.triajePeso), Replace(mo_DOAtencionesCENew.TriajePresion, "_", ""), Val(mo_DOAtencionesCENew.triajeTalla)

    End If
    If Val(mo_DOAtencionesCENew.triajePeso) >= 0 Then
        Me.UcDiagnosticoDetalle1.PesoKg = Val(mo_DOAtencionesCENew.triajePeso)
        lnPesoKg = Val(mo_DOAtencionesCENew.triajePeso)
    End If
    '===============================================
    'presion_lostfocus
'    If UcProgramaMaterno.Visible = True And Val(mo_DOAtencionesCENew.TriajePeso) > 0 And Val(mo_DOAtencionesCENew.TriajeTalla) > 0 Then
'       UcProgramaMaterno.Actualiza_Peso_Presion Val(mo_DOAtencionesCENew.TriajePeso), mo_DOAtencionesCENew.TriajePresion, Val(mo_DOAtencionesCENew.TriajeTalla)
'    End If
    '===============================================
    'talla lostfocus
'    If ucPerinatal1.Visible = True And Val(mo_DOAtencionesCENew.TriajeTalla) > 0 Then
'       ucPerinatal1.ActualizaGraficoYDiagnosticosAutomaticamente Val(mo_DOAtencionesCENew.TriajePeso), Val(mo_DOAtencionesCENew.TriajeTalla)
'    End If
'    If UcProgramaMaterno.Visible = True And Val(mo_DOAtencionesCENew.TriajePeso) > 0 And Val(mo_DOAtencionesCENew.TriajeTalla) > 0 Then
'       UcProgramaMaterno.Actualiza_Peso_Presion Val(mo_DOAtencionesCENew.TriajePeso), Replace(mo_DOAtencionesCENew.TriajePresion, "_", ""), Val(mo_DOAtencionesCENew.TriajeTalla)
'    End If
End Sub

Private Function RetornaObjetoDatosTriaje() As DOAtencionesCE
    Dim oDOAtencionesCE As DOAtencionesCE
    
    Set oDOAtencionesCE = ucTriajeVisorCE.DOAtencionCE
    
    If oDOAtencionesCE Is Nothing Then
        Set oDOAtencionesCE = New DOAtencionesCE
    End If
    Set RetornaObjetoDatosTriaje = oDOAtencionesCE
End Function

'mgaray20141013
'Private Function getDescripcionEdad(ml_lnEdadEnDias As Long, ml_lnIdTipoEdad As Long, _
'                md_fechaNacimiento As Date, md_fechaIngreso As Date) As String
'    Dim ls_edad As String
'    ls_edad = Trim(Str(ml_lnEdadEnDias)) & " " & SIGHEntidades.EdadDevuelveTipo(ml_lnIdTipoEdad)
'    If ml_lnIdTipoEdad = 1 Then
'        Dim oEdad As Edad
'        oEdad = calcularEdadDisgregada(md_fechaNacimiento, md_fechaIngreso)
'        If oEdad.EdadMes > 0 Then
'            ls_edad = ls_edad & ", " & oEdad.EdadMes & " " & SIGHEntidades.EdadDevuelveTipo(2)
'        End If
'    End If
'    getDescripcionEdad = ls_edad
'End Function

'Actualizado 27102014
Public Sub OcultarBotonesImpresionReceta(lbOcultar As Boolean)
    btnImprimir.Visible = lbOcultar
    btnImprimirOrden.Visible = lbOcultar
    If mi_Opcion = sghConsultar Then
       UcRecetas1.OcultarBotonesImpresionReceta True
    Else
       UcRecetas1.OcultarBotonesImpresionReceta lbOcultar
    End If
    If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE Then
        If lbOcultar = True Then
            mb_ControlNuevoMaterno = False
        Else
            mb_ControlNuevoMaterno = True
        End If
    End If
End Sub


Private Sub btnImprimir_Click()
    Dim ModificarDatos As Boolean
    If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE Then
        If ml_FechaReceta = 0 Then
           ml_FechaReceta = lcBuscaParametro.RetornaFechaHoraServidorSQL
        End If
        ModificarDatos = mo_AdminAdmision.RecetaModificar(mo_Atenciones.idCuentaAtencion, mo_Atenciones.IdServicioIngreso, ml_idUsuario, _
                                         lnRecetaRayosX, lnRecetaEcografiaO, lnRecetaEcografiaG, lnRecetaTomografia, _
                                         lnRecetaAnatomiaP, lnRecetaPatologiaC, lnRecetaBancoS, lnRecetaFarmacia, _
                                         Me.UcRecetas1.DevuelveRayosX, Me.UcRecetas1.DevuelveEcografiaO, _
                                         Me.UcRecetas1.DevuelveEcografiaG, Me.UcRecetas1.DevuelveTomografia, _
                                         Me.UcRecetas1.DevuelveAnatomia, Me.UcRecetas1.DevuelvePatologia, _
                                         Me.UcRecetas1.DevuelveBancoSangre, Me.UcRecetas1.DevuelveFarmacia, ml_FechaReceta, _
                                         mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "Paciente : " & ms_NombrePaciente, Me.idMedico, _
                                         False, Me.UcRecetas1.DevuelveOtrosCpt, lnRecetaOtrosCpt)
        If ModificarDatos = True Then
            If lnRecetaRayosX > 0 Or lnRecetaEcografiaO > 0 Or lnRecetaEcografiaG > 0 Or lnRecetaTomografia > 0 Or _
                     lnRecetaAnatomiaP > 0 Or lnRecetaPatologiaC > 0 Or lnRecetaBancoS > 0 Or lnRecetaFarmacia > 0 Then
'               Me.ucRecetas1.Tratamiento = Trim(TxtCitaTratamiento.Text)
'               Me.ucRecetas1.CargaNumeroDeRecetaEimprime lnRecetaRayosX, lnRecetaEcografiaO, lnRecetaEcografiaG, lnRecetaTomografia, _
'                                                         lnRecetaAnatomiaP, lnRecetaPatologiaC, lnRecetaBancoS, lnRecetaFarmacia, True
                Me.UcRecetas1.ImprimeOrdenMedica True
            End If
        End If
    End If
End Sub

Private Sub btnImprimirOrden_Click()
    Dim ModificarDatos As Boolean
    If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE Then
        If ml_FechaReceta = 0 Then
           ml_FechaReceta = lcBuscaParametro.RetornaFechaHoraServidorSQL
        End If
        ModificarDatos = mo_AdminAdmision.RecetaModificar(mo_Atenciones.idCuentaAtencion, mo_Atenciones.IdServicioIngreso, ml_idUsuario, _
                                         lnRecetaRayosX, lnRecetaEcografiaO, lnRecetaEcografiaG, lnRecetaTomografia, _
                                         lnRecetaAnatomiaP, lnRecetaPatologiaC, lnRecetaBancoS, lnRecetaFarmacia, _
                                         Me.UcRecetas1.DevuelveRayosX, Me.UcRecetas1.DevuelveEcografiaO, _
                                         Me.UcRecetas1.DevuelveEcografiaG, Me.UcRecetas1.DevuelveTomografia, _
                                         Me.UcRecetas1.DevuelveAnatomia, Me.UcRecetas1.DevuelvePatologia, _
                                         Me.UcRecetas1.DevuelveBancoSangre, Me.UcRecetas1.DevuelveFarmacia, ml_FechaReceta, _
                                         mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "Paciente : " & ms_NombrePaciente, Me.idMedico, _
                                         False, Me.UcRecetas1.DevuelveOtrosCpt, lnRecetaOtrosCpt)
        If ModificarDatos = True Then
            If lnRecetaRayosX > 0 Or lnRecetaEcografiaO > 0 Or lnRecetaEcografiaG > 0 Or lnRecetaTomografia > 0 Or _
                     lnRecetaAnatomiaP > 0 Or lnRecetaPatologiaC > 0 Or lnRecetaBancoS > 0 Or lnRecetaFarmacia > 0 Then
'               Me.ucRecetas1.Tratamiento = Trim(TxtCitaTratamiento.Text)
'               Me.ucRecetas1.CargaNumeroDeRecetaEimprime lnRecetaRayosX, lnRecetaEcografiaO, lnRecetaEcografiaG, lnRecetaTomografia, _
'                                                         lnRecetaAnatomiaP, lnRecetaPatologiaC, lnRecetaBancoS, lnRecetaFarmacia, True
                Me.UcRecetas1.ImprimeOrdenMedica False
            End If
        End If
    End If
End Sub

'mgaray20141024
Private Function getIdOrdenServicioInmunizaciones(lIdAtencion As Long) As Long
    Dim oRsOrdenServicio As Recordset
    Dim lOrdenServicio As Long
    
    lOrdenServicio = 0
    
    Set oRsOrdenServicio = mo_AdminAdmision.PerinatalBuscarOrdenServicioInmunizacion(lIdAtencion)
    
    ml_idOrdenServicioInmunizaciones = 0
    If oRsOrdenServicio.RecordCount > 0 Then
        If Not IsNull(oRsOrdenServicio.Fields!IdOrden) Then
        lOrdenServicio = oRsOrdenServicio.Fields!IdOrden
        End If
    End If
    getIdOrdenServicioInmunizaciones = lOrdenServicio
End Function

Private Function ocultarControles()
    Label34.Visible = False
    txtNroHijos.Visible = False
End Function

'mgaray201410f
Private Function BloqueoOpcionesMorbilidad() As Boolean
    If ucPerinatal1.Visible = True Or UcProgramaMaterno.Visible = True Then
        btnCpt.Visible = False
        UcDiagnosticoDetalle1.DeshabilitarEdicionDatos
'        TabDx.TabVisible(0) = False
    Else
        UcDiagnosticoDetalle1.HabilitarEdicionDatos
'        TabDx.TabVisible(0) = True
    End If
End Function

Private Function retornaRsProductoParaCpt() As ADODB.Recordset
    Dim mrs_FacturacionProductos As New ADODB.Recordset
    
    With mrs_FacturacionProductos
        .Fields.Append "IdFacturacionProducto", adInteger
        .Fields.Append "IdProducto", adInteger
        .Fields.Append "Codigo", adVarChar, 255, adFldIsNullable
        .Fields.Append "NombreProducto", adVarChar, 255, adFldIsNullable
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
        .Fields.Append "Grupo", adInteger
        .Fields.Append "SubGrupo", adInteger
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set retornaRsProductoParaCpt = mrs_FacturacionProductos
End Function

Sub cargaCptDesdeProgramasPerinatalMaterno(bEsPerinatal As Boolean, oRsProcedimientos As ADODB.Recordset, _
                    Optional lIdOrdenImunizaciones As Long = 0)
    If Me.ucPerinatal1.Visible = False And UcProgramaMaterno.Visible = False Then
       Exit Sub
    End If
    On Error GoTo errPer
    Dim oDOFactOrdenServicio As New DoFactOrdenServ
    Dim oDoCatalogoServicioHosp As New DOFinanciamientoCatalogoServ
    Dim mrs_FacturacionProductos As New Recordset
    Dim oRsVacunas As New Recordset
    Dim oConexion As New Connection
    Dim oRsTmp1 As New Recordset
    Dim oFactOrdenServicio As New FactOrdenServicio
    Dim lnPrecioUnitario As Double, lnIdPuntoCarga As Long
    Dim lnIdProducto As Long
    'mgaray201410e
'    Dim oRsProcedimientos As New Recordset
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion
    
    With oDOFactOrdenServicio
         .fechacreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL      'Now
         .idCuentaAtencion = Me.idCuentaAtencion
         .idestadofacturacion = sghEstadoFacturacion.sghAtendido
         .IdFuenteFinanciamiento = ml_IdFuenteFinanciamiento
         .idPaciente = ml_IdPaciente
         .idPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaServicioHospitalizacion  'consumo en el servicio
         .idTipoFinanciamiento = ml_IdFormaPago
         .idUsuario = ml_idUsuario
         .IdUsuarioAuditoria = ml_idUsuario
         .FechaDespacho = .fechacreacion
         .IdUsuarioDespacho = ml_idUsuario
         .FechaHoraRealizaCpt = .fechacreacion
    End With
    'mgaray20141024
    If oRsServiciosIntermedios.RecordCount > 0 Then
        Set oFactOrdenServicio.Conexion = oConexion
        oDOFactOrdenServicio.IdOrden = DevuelveIdOrdenDiferenteInmunizaciones(lIdOrdenImunizaciones)
        If oFactOrdenServicio.SeleccionarPorId(oDOFactOrdenServicio) = True Then
        End If
        Set oFactOrdenServicio.Conexion = Nothing
    End If
    
'    Set oRsProcedimientos = Me.ucPerinatal1.DevuelveCptFrecuentes
    
    If oRsProcedimientos.RecordCount > 0 Then
        'mgaray201410f
        Set mrs_FacturacionProductos = retornaRsProductoParaCpt()
        'mgaray201410e
        If oRsProcedimientos.RecordCount > 0 Then
            oRsProcedimientos.MoveFirst
            Do While Not oRsProcedimientos.EOF
                If bEsPerinatal = True Then
                    lnIdProducto = oRsProcedimientos!ID
                Else
                    lnIdProducto = oRsProcedimientos!idProducto
                End If
                If mo_ReglasComunes.ProcedimientoEsParaCpt(lnIdProducto) = True Then
                     Set oDoCatalogoServicioHosp = mo_AdminFacturacion.CatalogoServiciosHospSeleccionarPorId(lnIdProducto, _
                                                                                                             ml_IdFormaPago, _
                                                                                                             oConexion)
                     lnPrecioUnitario = oDoCatalogoServicioHosp.PrecioUnitario
                     mrs_FacturacionProductos.AddNew
                     mrs_FacturacionProductos.Fields!Codigo = ""
                     mrs_FacturacionProductos.Fields!idProducto = lnIdProducto
                     mrs_FacturacionProductos.Fields!NombreProducto = oRsProcedimientos!procedimiento
                     mrs_FacturacionProductos.Fields!labConfHIS = oRsProcedimientos!labConfHIS
                     mrs_FacturacionProductos.Fields!PrecioUnitario = lnPrecioUnitario
                     mrs_FacturacionProductos.Fields!TotalPorPagar = lnPrecioUnitario
                     mrs_FacturacionProductos.Fields!Cantidad = 1
                     mrs_FacturacionProductos.Fields!idestadofacturacion = 1
                     mrs_FacturacionProductos.Update
                End If
                oRsProcedimientos.MoveNext
            Loop
            If mrs_FacturacionProductos.RecordCount = 0 Then
                GoTo eliminarOrden
            End If
        
        End If
        'mgaray20141024
        If oDOFactOrdenServicio.IdOrden = 0 Or mi_Opcion = sghAgregar Then
            If mo_AdminFacturacion.FactOrdenServicioAgregar(oDOFactOrdenServicio, mrs_FacturacionProductos, _
                                                         mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Me.Caption, _
                                                         ml_IdServicio, 0, 0) = True Then
            End If
        Else
            If mo_AdminFacturacion.FactOrdenServicioModificar(oDOFactOrdenServicio, mrs_FacturacionProductos, _
                                                         mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Me.Caption) = True Then
            End If
        End If
    Else
eliminarOrden:
        If oRsServiciosIntermedios.RecordCount > 0 And oDOFactOrdenServicio.IdOrden > 0 And mi_Opcion <> sghAgregar Then
            If mo_AdminFacturacion.FactOrdenServicioEliminar(oDOFactOrdenServicio, mo_lnIdTablaLISTBARITEMS, _
                                                                          mo_lcNombrePc, Me.Caption, 0, 0) = True Then
            End If
        End If
    End If
    oConexion.Close
    Set oDOFactOrdenServicio = Nothing
    Set oDoCatalogoServicioHosp = Nothing
    Set mrs_FacturacionProductos = Nothing
    Set oRsVacunas = Nothing
    Set oRsTmp1 = Nothing
    Set oConexion = Nothing
    Set oFactOrdenServicio = Nothing
errPer:
End Sub

'mgaray201412a
Private Function DevuelveIdOrdenDiferenteInmunizaciones(lIdOrdenInmunizacion As Long) As Long
    Dim oRsServicios As ADODB.Recordset
    Dim lIdOrden As Long
    
    DevuelveIdOrdenDiferenteInmunizaciones = 0
    
    If oRsServiciosIntermedios.RecordCount > 0 Then
        Set oRsServicios = oRsServiciosIntermedios.Clone()
        oRsServicios.MoveFirst
        While oRsServicios.EOF = False
            If lIdOrdenInmunizacion <> oRsServicios!IdOrden Then
                DevuelveIdOrdenDiferenteInmunizaciones = oRsServicios!IdOrden
                Exit Function
            End If
            oRsServicios.MoveNext
        Wend
    End If
End Function


'***debb-27/05/2015
Private Sub btnQuitarCpt_Click()
    Dim oCpt As New FacOrdenServicioDetalle
    oCpt.FormMostradoDesde = 1
    oCpt.lbNOValidaCodigoPrestacion = True
    oCpt.PuntoCarga = 1   'consumo en el servicio
    oCpt.Opcion = sghEliminar
    oCpt.IdOrden = ml_idOrden
    oCpt.idUsuario = ml_idUsuario
    If ml_AScorrelativo = 0 Then
       oCpt.idCuentaAtencion = ml_idCuentaAtencion
    Else
       oCpt.idCuentaAtencion = ml_idOrden_idCuenta
    End If
    oCpt.Show 1
    Set oCpt = Nothing
'    If ml_AScorrelativo = 0 Then
'       CargaCPTrealizadosEnElServicio
'    Else
       CargaCPTrealizadosEnVariosServicios False
'    End If
End Sub
'***debb-27/05/2015
Private Sub grdOtrosCpt_Click()
    Dim rsRecordset As ADODB.Recordset
    Set rsRecordset = grdOtrosCpt.DataSource
    On Error Resume Next
    ml_idOrden = rsRecordset("IdOrden")
    ml_idOrden_idCuenta = rsRecordset("IdCuentaAtencion")
    Set rsRecordset = Nothing
End Sub
'***debb-14/05/2015
Private Sub grdOtrosCpt_AfterRowActivate()
    Dim rsRecordset As ADODB.Recordset
    Set rsRecordset = grdOtrosCpt.DataSource
    On Error Resume Next
    ml_idOrden = rsRecordset("IdOrden")
    ml_idOrden_idCuenta = rsRecordset("IdCuentaAtencion")
    Set rsRecordset = Nothing
End Sub

'debb-09/06/2016
Sub ActualizaVacunasDesdeModuloPerinatalAS()
    If Me.ucPerinatalAS1.Visible = False Then
       Exit Sub
    End If
    On Error GoTo errPer
    Dim oDOFactOrdenServicio As New DoFactOrdenServ
    Dim oDoCatalogoServicioHosp As New DOFinanciamientoCatalogoServ
    Dim mrs_FacturacionProductos As New Recordset
    Dim oRsVacunas As New Recordset
    Dim oConexion As New Connection
    Dim oRsTmp1 As New Recordset
    Dim oFactOrdenServicio As New FactOrdenServicio
    Dim lnPrecioUnitario As Double, lnIdPuntoCarga As Long
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion
    
    With oDOFactOrdenServicio
         .fechacreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL      'Now
         .idCuentaAtencion = Me.idCuentaAtencion
         .idestadofacturacion = sghEstadoFacturacion.sghAtendido
         .IdFuenteFinanciamiento = ml_IdFuenteFinanciamiento
         .idPaciente = ml_IdPaciente
         .idPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaServicioHospitalizacion   'consumo en el servicio
         .idTipoFinanciamiento = ml_IdFormaPago
         .idUsuario = ml_idUsuario
         .IdUsuarioAuditoria = ml_idUsuario
         .FechaDespacho = .fechacreacion
         .IdUsuarioDespacho = ml_idUsuario
         .FechaHoraRealizaCpt = .fechacreacion
    End With
    If oRsServiciosIntermedios.RecordCount > 0 Then
        oRsServiciosIntermedios.MoveFirst
        Set oFactOrdenServicio.Conexion = oConexion
        oDOFactOrdenServicio.IdOrden = oRsServiciosIntermedios!IdOrden
        If oFactOrdenServicio.SeleccionarPorId(oDOFactOrdenServicio) = True Then
        End If
        Set oFactOrdenServicio.Conexion = Nothing
    End If
    Set oRsVacunas = Me.ucPerinatalAS1.DevuelveCptInmunizaciones
    If oRsVacunas.RecordCount > 0 Then
        With mrs_FacturacionProductos
              .Fields.Append "IdFacturacionProducto", adInteger
              .Fields.Append "IdProducto", adInteger
              .Fields.Append "Codigo", adVarChar, 255, adFldIsNullable
              .Fields.Append "NombreProducto", adVarChar, 255, adFldIsNullable
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
              .Fields.Append "labConfHis", adVarChar, 3, adFldIsNullable
              .Fields.Append "Grupo", adInteger
              .Fields.Append "SubGrupo", adInteger
              .CursorType = adOpenDynamic
              .LockType = adLockOptimistic
              .Open
        End With
        oRsVacunas.MoveFirst
        Do While Not oRsVacunas.EOF
             Set oDoCatalogoServicioHosp = mo_AdminFacturacion.CatalogoServiciosHospSeleccionarPorId(oRsVacunas!ID, _
                                                                                                     ml_IdFormaPago, _
                                                                                                     oConexion)
             lnPrecioUnitario = oDoCatalogoServicioHosp.PrecioUnitario
             mrs_FacturacionProductos.AddNew
             mrs_FacturacionProductos.Fields!Codigo = ""
             mrs_FacturacionProductos.Fields!idProducto = oRsVacunas!ID
             mrs_FacturacionProductos.Fields!NombreProducto = oRsVacunas!procedimiento
             mrs_FacturacionProductos.Fields!PrecioUnitario = lnPrecioUnitario
             mrs_FacturacionProductos.Fields!TotalPorPagar = lnPrecioUnitario
             mrs_FacturacionProductos.Fields!Cantidad = 1
             mrs_FacturacionProductos.Fields!idestadofacturacion = 1
             mrs_FacturacionProductos.Fields!labConfHIS = Space(3)
             mrs_FacturacionProductos.Update
             oRsVacunas.MoveNext
        Loop
        If oRsServiciosIntermedios.RecordCount = 0 Then
            If mo_AdminFacturacion.FactOrdenServicioAgregar(oDOFactOrdenServicio, mrs_FacturacionProductos, _
                                                         mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Me.Caption, _
                                                         ml_IdServicio, 0, 0) = True Then
            End If
        Else
            If mo_AdminFacturacion.FactOrdenServicioModificar(oDOFactOrdenServicio, mrs_FacturacionProductos, _
                                                         mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Me.Caption) = True Then
            End If
        End If
    ElseIf oRsServiciosIntermedios.RecordCount > 0 Then
        If mo_AdminFacturacion.FactOrdenServicioEliminar(oDOFactOrdenServicio, mo_lnIdTablaLISTBARITEMS, _
                                                                      mo_lcNombrePc, Me.Caption, 0, 0) = True Then
        End If
    End If
    oConexion.Close
    Set oDOFactOrdenServicio = Nothing
    Set oDoCatalogoServicioHosp = Nothing
    Set mrs_FacturacionProductos = Nothing
    Set oRsVacunas = Nothing
    Set oRsTmp1 = Nothing
    Set oConexion = Nothing
    Set oFactOrdenServicio = Nothing
errPer:
Exit Sub
Resume
End Sub
'debb-09/06/2016
Sub CargarDatosPerinatalAS(oConexion As Connection)
    If Me.ucPerinatalAS1.Visible = True Then
       TabDx.Tab = 1
       Dim lnEdadEnDias As Integer, lnIdTipoEdad As Integer
       Dim oDOAtencionesCE As DOAtencionesCE
       
       Set oDOAtencionesCE = RetornaObjetoDatosTriaje()
       lnEdadEnDias = ml_lnEdadEnDias
       lnIdTipoEdad = ml_lnIdTipoEdad
       'If lbCargaUnaSolaVez = False Then
           Me.ucPerinatalAS1.idUsuario = ml_idUsuario
           Me.ucPerinatalAS1.FechaAtencion = mo_Atenciones.FechaIngreso
           Me.ucPerinatalAS1.FechaNacimiento = mo_paciente.FechaNacimiento
           Me.ucPerinatalAS1.Inicializar
       'End If
       Me.ucPerinatalAS1.idPaciente = mo_Atenciones.idPaciente
       Me.ucPerinatalAS1.idAtencion = mo_Atenciones.idAtencion
       Me.ucPerinatalAS1.idTipoSexo = mo_paciente.idTipoSexo
       Me.ucPerinatalAS1.EdadEnMeses = sighEntidades.DevuelveEdadEnMeses(mo_paciente.FechaNacimiento, mo_Atenciones.FechaIngreso)
       'Me.ucPerinatalAS1.FechaAtencion = mo_Atenciones.FechaIngreso
       'Me.ucPerinatalAS1.FechaNacimiento = mo_paciente.FechaNacimiento
       Me.ucPerinatalAS1.CargaDatosAcontroles lnEdadEnDias, lnIdTipoEdad, Val(oDOAtencionesCE.triajePeso), _
                                            Val(oDOAtencionesCE.triajeTalla), oConexion
       SendKeys "{tab}"
    End If
End Sub
'debb-09/06/2016
Function GrabaAtencionPerinatalAS() As Boolean
    If ucPerinatalAS1.Visible = True Then
        Dim oDoPerinatalAtencion As New DoPerinatalAtencion
        Select Case mi_Opcion
        Case sghAgregar
             GrabaAtencionPerinatalAS = True
        Case sghModificar
             Set oDoPerinatalAtencion = Me.ucPerinatalAS1.DevuelveDatosGenerales
             oDoPerinatalAtencion.idPaciente = mo_Atenciones.idPaciente
             oDoPerinatalAtencion.IdUsuarioAuditoria = mo_Atenciones.IdUsuarioAuditoria
             If oDoPerinatalAtencion.idPerinatalAtencion = 0 Then
                GrabaAtencionPerinatalAS = mo_AdminAdmision.PerinatalCEAgregar(oDoPerinatalAtencion, _
                                                          Me.ucPerinatalAS1.DevuelveCptInmunizaciones, _
                                                          Me.ucPerinatalAS1.DevuelveCptFrecuentes, _
                                                          Me.ucPerinatalAS1.DevuelveDxDesarrollo, _
                                                          Me.ucPerinatalAS1.DevuelveDxMorbilidad, _
                                                          Me.ucPerinatalAS1.DevuelveMedicamentos, _
                                                          Me.ucPerinatalAS1.DevuelveDatosCred(False), _
                                                          Me.ucPerinatalAS1.DevuelvePerinatalAtencionCred1, _
                                                          mo_Atenciones.idAtencion)
             Else
                GrabaAtencionPerinatalAS = mo_AdminAdmision.PerinatalCEModificar(oDoPerinatalAtencion, _
                                                          Me.ucPerinatalAS1.DevuelveCptInmunizaciones, _
                                                          Me.ucPerinatalAS1.DevuelveCptFrecuentes, _
                                                          Me.ucPerinatalAS1.DevuelveDxDesarrollo, _
                                                          Me.ucPerinatalAS1.DevuelveDxMorbilidad, _
                                                          Me.ucPerinatalAS1.DevuelveMedicamentos, _
                                                          Me.ucPerinatalAS1.DevuelveDatosCred(False), _
                                                          Me.ucPerinatalAS1.DevuelvePerinatalAtencionCred1, _
                                                          mo_Atenciones.idAtencion)
             End If
        Case sghEliminar
            oDoPerinatalAtencion.idPaciente = mo_Atenciones.idPaciente
            GrabaAtencionPerinatalAS = mo_AdminAdmision.PerinatalCEeliminar(oDoPerinatalAtencion, mo_Atenciones.idAtencion)
        End Select
        If GrabaAtencionPerinatalAS = False Then
           MsgBox "Falló al Grabar PERINATAL" & Chr(13) & mo_AdminAdmision.MensajeError
        End If
        Set oDoPerinatalAtencion = Nothing
    End If
End Function


Private Sub txtFextension_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFextension

End Sub
Private Sub txtNroReferenciaDestino_KeyDown(KeyCode As Integer, Shift As Integer)
       mo_Teclado.RealizarNavegacion KeyCode, txtNroReferenciaDestino

End Sub
Private Sub cmbServicioReferenciaD_KeyDown(KeyCode As Integer, Shift As Integer)
           mo_Teclado.RealizarNavegacion KeyCode, cmbServicioReferenciaD

End Sub
Private Sub txtFtramite_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtFtramite
End Sub

Sub VerificaPermisos()
  lbTienePermisoParaRegistrarAtencionesPasadas = mo_AdminAdmision.TienePermisosParaModificarAtencionesPasadas
    
    lbTienePermisoParaImprimirAtencion = False
    Dim ms_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
    Dim oRsPermisosTabs As New Recordset
    Set oRsPermisosTabs = ms_ReglasSeguridad.UsuariosRolesSeleccionarPermisosTodos(ml_idUsuario)
    If oRsPermisosTabs.RecordCount > 0 Then
       Do While Not oRsPermisosTabs.EOF
          Select Case oRsPermisosTabs.Fields!IdPermiso
          Case 411    'Admisión CE - Imprimir Atencion
               lbTienePermisoParaImprimirAtencion = True
          End Select
          oRsPermisosTabs.MoveNext
       Loop
    End If
    Set oRsPermisosTabs = Nothing
    Set ms_ReglasSeguridad = Nothing
End Sub


Sub ActividadesHIS(oRsItemsElegidos As Recordset)
'        Dim oRsActividades As New Recordset
        Dim oRsTmp1 As New Recordset, oRsTmp2 As New Recordset, oRsDx As New Recordset
       Dim lcNombre As String, lnGrupo As Integer, lnSubGrupo As Integer, lbPrimerReg As Boolean
'        Dim lnEdad As Long
        Dim lcEligio As Boolean, lcEligioLab As String, lnEligioTipo As Integer, lnEligioUPS As Long
'        Dim ln_IdCuentaAtencion As Long, ln_IdOrden As Long, ln_Fua As Integer, lc_Consultorio As String
'        Dim ln_idServicio As Long, lc_FuaCodigoPrestacion As String, lbUnaSolaVez As Boolean
        Dim lc_id As String, lnPrecioUnitario As Double, ln_idServicioPaciente As Long
        Dim oFactOrdenServicio As New FactOrdenServicio
        Dim oDOFactOrdenServicio As New DoFactOrdenServ
        Dim mrs_FacturacionProductos As New Recordset
        Dim oDoCatalogoServicioHosp As New DOFinanciamientoCatalogoServ
'        With oRsActividades
'              .Fields.Append "GrupoTIT", adVarChar, 3, adFldIsNullable
'              .Fields.Append "Grupo", adInteger
'              .Fields.Append "SubGrupo", adInteger
'              .Fields.Append "lab", adVarChar, 3, adFldIsNullable
'              .Fields.Append "Tipo", adVarChar, 20, adFldIsNullable
'              If lnOpcion = 2 Then
'              End If
'              .Fields.Append "id", adVarChar, 20, adFldIsNullable
'              .Fields.Append "Nombre", adVarChar, 255, adFldIsNullable
'              .Fields.Append "Elija", adBoolean
'              .Fields.Append "ElijaTipo", adInteger
'              .Fields.Append "ElijaUPS", adInteger
'              .Fields.Append "ElijaLab", adVarChar, 3, adFldIsNullable
'              .Fields.Append "IdCuentaAtencion", adInteger, 4, adFldIsNullable
'              .Fields.Append "IdOrden", adInteger, 4, adFldIsNullable
'              .Fields.Append "Fua", adInteger
'              .Fields.Append "Consultorio", adVarChar, 100, adFldIsNullable
'              .Fields.Append "IdServicio", adInteger
'              .Fields.Append "FuaCodigoPrestacion", adVarChar, 3, adFldIsNullable
'              .Fields.Append "idTipo", adInteger
'              .Fields.Append "idServicioPaciente", adInteger
'              .CursorType = adOpenKeyset
'              .LockType = adLockOptimistic
'              .Open
'        End With
'        If lnOpcion = 2 Then
'           '********************* con actividades vacias ******************
'        Else
'           '********************* con actividades ya registradas ******************
'        Set oRsTmp1 = mo_AdminAdmision.ServiciosAtenSimultaneaImpHISxUPS(mo_AdminAdmision.BuscaUPSactualDelPaciente(mo_Atenciones.IdServicioIngreso))
'
'        Dim oEdad As Edad, lbContiuar9 As Boolean
'        '
'       ' oEdad = calcularEdadDisgregada(mo_paciente.FechaNacimiento, mo_Atenciones.FechaIngreso)
'        Select Case cmbTipoEdad1.ListIndex
'        Case 0   'año
'             oEdad.EdadAnio = Val(txtEdad1.Text)
'        Case 1   'meses
'             oEdad.EdadMes = Val(txtEdad1.Text)
'        Case Else
'             oEdad.EdadDia = Val(txtEdad1.Text)
'        End Select
'        oEdad.TipoEdad = cmbTipoEdad1.ListIndex + 1
'        '
'        If oEdad.EdadAnio > 0 Then
'           oRsTmp1.Filter = "idtipoedad=1 and edadinicio" & IIf(ml_ups = "301202", "=", "<=") & oEdad.EdadAnio
'        ElseIf oEdad.EdadMes > 0 Then
'           oRsTmp1.Filter = "idtipoedad=2 and edadinicio=" & oEdad.EdadMes
'        Else
'          oRsTmp1.Filter = "idtipoedad=3"
'        End If
'        If oRsTmp1.RecordCount > 0 Then
'           Set oRsDx = Me.UcDiagnosticoDetalle1.DevuelveDx
'           oRsTmp1.MoveFirst
'           Do While Not oRsTmp1.EOF
'              lnGrupo = oRsTmp1!Grupo
'
'              lbPrimerReg = True
'              Do While Not oRsTmp1.EOF And lnGrupo = oRsTmp1!Grupo
'                 lbContiuar9 = True
'                 If oEdad.EdadDia > 0 And oEdad.EdadAnio = 0 And oEdad.EdadMes = 0 Then
'                    If Not (oEdad.EdadDia >= oRsTmp1!EdadInicio And oEdad.EdadDia <= oRsTmp1!EdadFinal) Then
'                       lbContiuar9 = False
'                    End If
'                 End If
'
'                 If (lnPesoKg >= oRsTmp1!PesoKgMenor And lnPesoKg <= oRsTmp1!PesoKgMayor) And lbContiuar9 = True Then
'                    lnSubGrupo = oRsTmp1!subgrupoOrden
'                    lcNombre = ""
'                    Select Case oRsTmp1!idTipo
'                    Case 1  'cpt
'                            Set oRsTmp2 = mo_AdminCaja.FactCatalogoServiciosSeleccionarPorCodigoOnombre(oRsTmp1!cpt_dx, "")
'                            If oRsTmp2.RecordCount > 0 Then
'                               lcNombre = Left(oRsTmp2!Nombre, 255)
'                            End If
'                            oRsTmp2.Close
'                    Case 3  'dx
'                            Set oRsTmp2 = mo_AdminServiciosComunes.DiagnosticosSeleccionarXCodigo(oRsTmp1!cpt_dx)
'                            If oRsTmp2.RecordCount > 0 Then
'                               lcNombre = Left(oRsTmp2!Descripcion, 255)
'                            End If
'                             oRsTmp2.Close
'                    End Select
'                    '
'                    lcEligio = False
'                    lcEligioLab = ""
'                    lnEligioTipo = 102
'                    lnEligioUPS = ml_idCuentaAtencion    'mo_Atenciones.IdServicioIngreso
'                    ln_IdCuentaAtencion = ml_idCuentaAtencion: ln_IdOrden = 0: ln_Fua = 0: lc_Consultorio = ml_lcServicio
'                    ln_idServicio = ml_idCuentaAtencion: lc_FuaCodigoPrestacion = "": ln_idServicioPaciente = mo_Atenciones.IdServicioIngreso
'                    If oRsTmp1!idTipo = 1 Then
'                        oRsGrdOtrosCpt.Filter = "grupo=" & lnGrupo & " and subgrupo=" & lnSubGrupo & _
'                                                " and codigo='" & Trim(oRsTmp1!cpt_dx) & "'"
'                        If oRsGrdOtrosCpt.RecordCount > 0 Then
'                            lcEligio = True
'                            lcEligioLab = oRsGrdOtrosCpt!labConfHIS
'                            lnEligioTipo = IIf(IsNull(oRsGrdOtrosCpt!idTipoDiagnostico), 102, oRsGrdOtrosCpt!idTipoDiagnostico)
'                            lnEligioUPS = oRsGrdOtrosCpt!idCuentaAtencion
'                            ln_IdCuentaAtencion = oRsGrdOtrosCpt!idCuentaAtencion
'                            ln_IdOrden = oRsGrdOtrosCpt!IdOrden
'                            ln_Fua = oRsGrdOtrosCpt!fua
'                            lc_Consultorio = oRsGrdOtrosCpt!Consultorio
'                            ln_idServicio = oRsGrdOtrosCpt!IdServicio
'                            ln_idServicioPaciente = oRsGrdOtrosCpt!IdServicio
'                        End If
'                    Else
'                        oRsDx.Filter = "grupo=" & lnGrupo & " and subgrupo=" & lnSubGrupo & _
'                                       " and CodigoCIE2004='" & Trim(oRsTmp1!cpt_dx) & "'"
'                        If oRsDx.RecordCount > 0 Then
'                            lcEligio = True
'                            lcEligioLab = IIf(IsNull(oRsDx!labConfHIS), "", oRsDx!labConfHIS)
'                            lnEligioTipo = oRsDx!idTipoDiagnostico
'                            lnEligioUPS = oRsDx!idCuentaAtencion
'                            ln_IdCuentaAtencion = oRsDx!idCuentaAtencion
'                            ln_Fua = oRsDx!fua
'                            lc_Consultorio = oRsDx!Consultorio
'                            ln_idServicio = oRsDx!idCuentaAtencion
'                            lc_FuaCodigoPrestacion = IIf(IsNull(oRsDx!FuaCodigoPrestacion), "", oRsDx!FuaCodigoPrestacion)
'                            ln_idServicioPaciente = oRsDx!IdServicio
'                        End If
'
'                    End If
'                    '
'                    oRsActividades.AddNew
'                    If lbPrimerReg = True Then
'                       lbPrimerReg = False
'                       oRsActividades!GrupoTIT = Trim(Str(lnGrupo))
'                    Else
'                       oRsActividades!GrupoTIT = ""
'                    End If
'                    oRsActividades!Grupo = lnGrupo
'                    oRsActividades!subgrupo = lnSubGrupo
'                    oRsActividades!Lab = IIf(IsNull(oRsTmp1!Lab), " ", oRsTmp1!Lab)
'                    oRsActividades!Id = oRsTmp1!cpt_dx
'                    oRsActividades!tipo = oRsTmp1!dTipo
'                    oRsActividades!Nombre = lcNombre
'                    oRsActividades!Elija = lcEligio
'                    oRsActividades!elijaTipo = lnEligioTipo - 100
'                    oRsActividades!elijaUPS = lnEligioUPS
'                    oRsActividades!elijaLAB = lcEligioLab
'                    oRsActividades!idCuentaAtencion = ln_IdCuentaAtencion
'                    oRsActividades!IdOrden = ln_IdOrden
'                    oRsActividades!fua = ln_Fua
'                    oRsActividades!Consultorio = lc_Consultorio
'                    oRsActividades!IdServicio = ln_idServicio
'                    oRsActividades!FuaCodigoPrestacion = lc_FuaCodigoPrestacion
'                    oRsActividades!idTipo = oRsTmp1!idTipo
'                    oRsActividades!IdServicioPaciente = ln_idServicioPaciente
'                    oRsActividades.Update
'                 End If
'                 oRsTmp1.MoveNext
'                 If oRsTmp1.EOF Then
'                    Exit Do
'                 End If
'              Loop
'           Loop
'
'           oRsGrdOtrosCpt.Filter = ""
'           If oRsGrdOtrosCpt.RecordCount > 0 Then
'              oRsGrdOtrosCpt.MoveFirst
'           End If
'           oRsDx.Filter = ""
'           If oRsDx.RecordCount > 0 Then
'              oRsDx.MoveFirst
'           End If
'        End If
'        oRsTmp1.Close
'        If oRsActividades.RecordCount > 0 Then
'            Dim oAdmisionCEatencSimultanea As New AdmisionCEatencSimultanea
'            Dim oRsItemsElegidos As New Recordset
'            oAdmisionCEatencSimultanea.FormLlamante = "ACTIVIDADES"
'            Set oAdmisionCEatencSimultanea.oRsFua = oRsActividades
'            Set oAdmisionCEatencSimultanea.oRsItemsElegidos = oRsTipoDx
'            oAdmisionCEatencSimultanea.Show 1
'            If oAdmisionCEatencSimultanea.idCuentaAtencion = 1 Then
'               Set oRsItemsElegidos = oAdmisionCEatencSimultanea.ItemsMasivosElegidos

               'Dx
               Me.UcDiagnosticoDetalle1.EliminaLosQueTienenGrupo
               If oRsItemsElegidos.State = 1 Then
                    If Not oRsItemsElegidos Is Nothing Then
                      oRsItemsElegidos.Filter = "idTipo=3"
                      Set Me.UcDiagnosticoDetalle1.oRsItemsElegidos = oRsItemsElegidos
                    End If
               End If
               'cpt
               With oDOFactOrdenServicio
                     .fechacreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL      'Now
                    ' .idCuentaAtencion = ml_idCuentaAtencion
                     '.idestadofacturacion = sghEstadoFacturacion.sghAnulado
                     .IdFuenteFinanciamiento = mo_Atenciones.IdFuenteFinanciamiento
                     .idPaciente = mo_Atenciones.idPaciente
                     .idPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaServicioHospitalizacion
                     .idTipoFinanciamiento = mo_Atenciones.IdFormaPago
                     .idUsuario = ml_idUsuario
                     .IdUsuarioAuditoria = ml_idUsuario
                     .FechaDespacho = .fechacreacion
                     .IdUsuarioDespacho = ml_idUsuario
                     .FechaHoraRealizaCpt = .fechacreacion
               End With
               If oRsGrdOtrosCpt.RecordCount > 0 Then
                  oRsGrdOtrosCpt.MoveFirst
                  Do While Not oRsGrdOtrosCpt.EOF
                     If oRsGrdOtrosCpt!Grupo > 0 Then
                        oDOFactOrdenServicio.IdOrden = oRsGrdOtrosCpt!IdOrden
                        If mo_AdminFacturacion.FactOrdenServicioEliminar(oDOFactOrdenServicio, mo_lnIdTablaLISTBARITEMS, _
                                                                                     mo_lcNombrePc, Me.Caption, 0, 0) = True Then
                        End If
                     End If
                     oRsGrdOtrosCpt.MoveNext
                  Loop
               End If
               If oRsItemsElegidos.State = 1 Then
                    If Not oRsItemsElegidos Is Nothing Then
                        oRsItemsElegidos.Filter = "idTipo=1"
                        If oRsItemsElegidos.RecordCount > 0 Then
                           oRsItemsElegidos.Sort = "ElijaUPS,grupo,id"
                           oRsItemsElegidos.MoveFirst
                           Do While Not oRsItemsElegidos.EOF
                                 oDOFactOrdenServicio.idCuentaAtencion = oRsItemsElegidos!idCuentaAtencion
                                 oDOFactOrdenServicio.idestadofacturacion = sghEstadoFacturacion.sghAtendido
                                 lnEligioUPS = oRsItemsElegidos!ElijaUPS
                                 lc_id = oRsItemsElegidos!ID
                                 ln_idServicioPaciente = oRsItemsElegidos!idServicioPaciente
                                 lnGrupo = oRsItemsElegidos!Grupo
                                 Set mrs_FacturacionProductos = retornaRsProductoParaCpt()
                                 Do While Not oRsItemsElegidos.EOF And _
                                          lc_id = oRsItemsElegidos!ID And _
                                          lnEligioUPS = oRsItemsElegidos!ElijaUPS And _
                                          lnGrupo = oRsItemsElegidos!Grupo
                                     Set oRsTmp1 = mo_AdminCaja.FactCatalogoServiciosSeleccionarPorCodigoOnombre(oRsItemsElegidos!ID, "")
                                     lnPrecioUnitario = 0
                                     Set oDoCatalogoServicioHosp = mo_AdminFacturacion.CatalogoServiciosHospSeleccionarPorId(oRsTmp1!idProducto, _
                                                                                              mo_Atenciones.IdFormaPago)
                                     If Not (oDoCatalogoServicioHosp Is Nothing) Then
                                        lnPrecioUnitario = oDoCatalogoServicioHosp.PrecioUnitario
                                     End If
                                     mrs_FacturacionProductos.AddNew
                                     mrs_FacturacionProductos.Fields!Codigo = ""
                                     mrs_FacturacionProductos.Fields!idProducto = oRsTmp1!idProducto
                                     mrs_FacturacionProductos.Fields!NombreProducto = oRsItemsElegidos!nombre
                                     mrs_FacturacionProductos.Fields!PrecioUnitario = lnPrecioUnitario
                                     mrs_FacturacionProductos.Fields!TotalPorPagar = lnPrecioUnitario
                                     mrs_FacturacionProductos.Fields!Cantidad = 1
                                     mrs_FacturacionProductos.Fields!idestadofacturacion = 1
                                     mrs_FacturacionProductos.Fields!Grupo = oRsItemsElegidos!Grupo
                                     mrs_FacturacionProductos.Fields!SubGrupo = oRsItemsElegidos!SubGrupo
                                     mrs_FacturacionProductos.Fields!labConfHIS = oRsItemsElegidos!ElijaLab
                                     mrs_FacturacionProductos.Update
                                     oRsItemsElegidos.MoveNext
                                     If oRsItemsElegidos.EOF Then
                                        Exit Do
                                     End If
                                 Loop
                                 If mo_AdminFacturacion.FactOrdenServicioAgregar(oDOFactOrdenServicio, _
                                                        mrs_FacturacionProductos, mo_lnIdTablaLISTBARITEMS, _
                                                        mo_lcNombrePc, Me.Caption, ln_idServicioPaciente, 0, 0) = True Then
                                 End If
                           Loop
                           
                        
                        End If
                    End If
               End If
'            End If
            Set oRsItemsElegidos = Nothing
'        End If
        Set oRsTmp1 = Nothing
        Set oRsDx = Nothing
        Set oFactOrdenServicio = Nothing
        Set oDOFactOrdenServicio = Nothing
        Set mrs_FacturacionProductos = Nothing
        Set oDoCatalogoServicioHosp = Nothing
         CargaCPTrealizadosEnVariosServicios False

End Sub
