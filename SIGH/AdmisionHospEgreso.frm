VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form AdmisionHospEgreso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Egresos"
   ClientHeight    =   9060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12105
   ControlBox      =   0   'False
   Icon            =   "AdmisionHospEgreso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   12105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   885
      Left            =   30
      TabIndex        =   10
      Top             =   8190
      Width           =   12015
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
         Height          =   675
         Left            =   10755
         Picture         =   "AdmisionHospEgreso.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   150
         Width           =   1185
      End
      Begin VB.CommandButton btnImprimeFichaSIS 
         Caption         =   "Imp.FUA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   9420
         Picture         =   "AdmisionHospEgreso.frx":1254
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   150
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar"
         DisabledPicture =   "AdmisionHospEgreso.frx":172D
         DownPicture     =   "AdmisionHospEgreso.frx":1BF1
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   6135
         Picture         =   "AdmisionHospEgreso.frx":20DD
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   150
         Width           =   1185
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "AdmisionHospEgreso.frx":25C9
         DownPicture     =   "AdmisionHospEgreso.frx":2A29
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   4560
         Picture         =   "AdmisionHospEgreso.frx":2E9E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   150
         Width           =   1185
      End
   End
   Begin TabDlg.SSTab TabEgresos 
      Height          =   8205
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   12045
      _ExtentX        =   21246
      _ExtentY        =   14473
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      ForeColor       =   13653559
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "3.1 Egreso"
      TabPicture(0)   =   "AdmisionHospEgreso.frx":3313
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "UcEpisodioClinico1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ucDiagnosticosEgreso"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraDatosReferenciaDestino"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraSoloEme"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "3.2 Complicaciones"
      TabPicture(1)   =   "AdmisionHospEgreso.frx":332F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ucDiagnosticoComplicaciones"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "3.3 Nacimientos"
      TabPicture(2)   =   "AdmisionHospEgreso.frx":334B
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(1)=   "ucNacimientoDetalle1"
      Tab(2).Control(2)=   "ucDiagnosticoNacimiento"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "3.4 Mortalidad"
      TabPicture(3)   =   "AdmisionHospEgreso.frx":3367
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame3"
      Tab(3).Control(1)=   "ucDiagnosticosMortalidad"
      Tab(3).ControlCount=   2
      Begin VB.Frame fraSoloEme 
         Height          =   1275
         Left            =   5790
         TabIndex        =   64
         Top             =   3120
         Visible         =   0   'False
         Width           =   6165
         Begin VB.ComboBox cmbEstadoLlegada 
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
            ItemData        =   "AdmisionHospEgreso.frx":3383
            Left            =   1035
            List            =   "AdmisionHospEgreso.frx":3390
            TabIndex        =   67
            Top             =   495
            Width           =   1920
         End
         Begin VB.ComboBox cmbComoLlego 
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
            ItemData        =   "AdmisionHospEgreso.frx":33B6
            Left            =   1035
            List            =   "AdmisionHospEgreso.frx":33C6
            TabIndex        =   65
            Top             =   150
            Width           =   1920
         End
         Begin VB.ComboBox cmbTipoAtencion 
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
            ItemData        =   "AdmisionHospEgreso.frx":33F6
            Left            =   4245
            List            =   "AdmisionHospEgreso.frx":3400
            TabIndex        =   66
            Top             =   165
            Width           =   1860
         End
         Begin VB.ComboBox cmbIdTipoGravedad 
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
            ItemData        =   "AdmisionHospEgreso.frx":341A
            Left            =   1035
            List            =   "AdmisionHospEgreso.frx":341C
            TabIndex        =   68
            Text            =   "cmbIdTipoGravedad"
            Top             =   855
            Width           =   2445
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
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
            Height          =   210
            Left            =   105
            TabIndex        =   72
            Top             =   555
            Width           =   555
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Como llegó"
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
            TabIndex        =   71
            Top             =   210
            Width           =   900
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Atención"
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
            Left            =   3105
            TabIndex        =   70
            Top             =   195
            Width           =   1155
         End
         Begin VB.Label lblGravedad 
            AutoSize        =   -1  'True
            Caption         =   "Gravedad"
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
            TabIndex        =   69
            Top             =   900
            Width           =   765
         End
      End
      Begin VB.Frame fraDatosReferenciaDestino 
         Caption         =   "Referencia Destino"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2145
         Left            =   105
         TabIndex        =   31
         Top             =   2190
         Width           =   5625
         Begin VB.CheckBox chkEnPDF 
            Caption         =   "En PDF"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   540
            TabIndex        =   74
            Top             =   1635
            Visible         =   0   'False
            Width           =   900
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
            Left            =   105
            Picture         =   "AdmisionHospEgreso.frx":341E
            Style           =   1  'Graphical
            TabIndex        =   60
            ToolTipText     =   "Imprimir recetas de farmacia"
            Top             =   1590
            Width           =   405
         End
         Begin VB.TextBox txtIdEstablecimientoDestino 
            Height          =   315
            Left            =   1590
            TabIndex        =   39
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox lblNombreDestinoReferencia 
            Height          =   315
            Left            =   2955
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   600
            Width           =   2580
         End
         Begin VB.ComboBox cmbIdTipoReferenciaDestino 
            Height          =   315
            Left            =   1590
            TabIndex        =   37
            Top             =   240
            Width           =   1635
         End
         Begin VB.CommandButton btnBuscarEstablecimientoDestino 
            Caption         =   "..."
            Height          =   315
            Left            =   2610
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   600
            Width           =   315
         End
         Begin VB.TextBox txtReferenciaD 
            Height          =   315
            Left            =   4530
            TabIndex        =   32
            Top             =   240
            Width           =   975
         End
         Begin PVCOMBOLibCtl.PVComboBox cmbServicioReferenciaD 
            Height          =   330
            Left            =   1590
            TabIndex        =   33
            Top             =   930
            Width           =   3945
            _Version        =   524288
            _cx             =   6959
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
            Left            =   1590
            TabIndex        =   34
            Top             =   1260
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
            Left            =   4185
            TabIndex        =   35
            Top             =   1260
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
         Begin VB.Label lblIdAtencion 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
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
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   5325
            TabIndex        =   73
            Top             =   1785
            Visible         =   0   'False
            Width           =   180
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
            Height          =   210
            Left            =   90
            TabIndex        =   54
            Top             =   300
            Width           =   1230
         End
         Begin VB.Label lblIdEstablecimientoDestino 
            BackStyle       =   0  'Transparent
            Caption         =   "Establec.Referen"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   90
            TabIndex        =   53
            Top             =   630
            Width           =   1410
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
            Left            =   90
            TabIndex        =   52
            Top             =   960
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
            Left            =   90
            TabIndex        =   51
            Top             =   1290
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
            Left            =   3255
            TabIndex        =   50
            Top             =   1290
            Width           =   840
         End
         Begin VB.Label lblNreferencia0 
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
            Left            =   3375
            TabIndex        =   40
            Top             =   285
            Width           =   1125
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Egreso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1725
         Left            =   120
         TabIndex        =   19
         Top             =   450
         Width           =   11835
         Begin VB.CommandButton cmdBuscaCamaEgreso 
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
            Left            =   7860
            Picture         =   "AdmisionHospEgreso.frx":38F7
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   540
            Width           =   360
         End
         Begin VB.CommandButton btnBuscarMedicosEgreso 
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
            Left            =   4275
            Picture         =   "AdmisionHospEgreso.frx":3E81
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   1275
            Width           =   360
         End
         Begin VB.ComboBox cmbServicioDestino 
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
            Left            =   7050
            TabIndex        =   6
            Top             =   1290
            Visible         =   0   'False
            Width           =   4380
         End
         Begin VB.TextBox txtNroCamaEgreso 
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
            Left            =   7050
            TabIndex        =   23
            Top             =   570
            Width           =   780
         End
         Begin VB.TextBox txtIdServicioEgreso 
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
            Left            =   1620
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   570
            Width           =   885
         End
         Begin VB.TextBox txtIdMedicoEgreso 
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
            Left            =   4650
            TabIndex        =   21
            Top             =   1290
            Width           =   885
         End
         Begin VB.TextBox lblNombreServicioEgreso 
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
            Left            =   2580
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   570
            Width           =   2970
         End
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
            Left            =   1620
            TabIndex        =   0
            Top             =   210
            Width           =   3960
         End
         Begin VB.ComboBox cmbCondicionAlta 
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
            Left            =   7050
            TabIndex        =   5
            Top             =   930
            Width           =   4380
         End
         Begin VB.ComboBox cmbTipoAlta 
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
            Left            =   7050
            TabIndex        =   4
            Top             =   210
            Width           =   4380
         End
         Begin VB.TextBox lblNombreMedicoEgreso 
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
            Left            =   1620
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   1290
            Width           =   2640
         End
         Begin MSMask.MaskEdBox txtHoraEgreso 
            Height          =   315
            Left            =   3045
            TabIndex        =   2
            Top             =   930
            Width           =   735
            _ExtentX        =   1296
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
         Begin MSMask.MaskEdBox txtFechaEgreso 
            Height          =   315
            Left            =   1620
            TabIndex        =   1
            Top             =   930
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
         Begin MSMask.MaskEdBox txtHoraEgresoAdm 
            Height          =   315
            Left            =   10665
            TabIndex        =   56
            Top             =   570
            Width           =   735
            _ExtentX        =   1296
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
         Begin MSMask.MaskEdBox txtFechaEgresoAdm 
            Height          =   315
            Left            =   9240
            TabIndex        =   57
            Top             =   570
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
         Begin VB.Label lblServicioDestino 
            AutoSize        =   -1  'True
            Caption         =   "Servicio Dest"
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
            Left            =   5940
            TabIndex        =   59
            Top             =   1350
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.Label lblFechaEgresoAdm 
            AutoSize        =   -1  'True
            Caption         =   "F.alta adm"
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
            Left            =   8415
            TabIndex        =   58
            Top             =   630
            Width           =   840
         End
         Begin VB.Label lblNroCamaEgreso 
            AutoSize        =   -1  'True
            Caption         =   "Cama egreso"
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
            Left            =   5940
            TabIndex        =   30
            Top             =   645
            Width           =   1050
         End
         Begin VB.Label Label49 
            Caption         =   "Servicio egreso"
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
            TabIndex        =   29
            Top             =   600
            Width           =   1395
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "Tipo alta"
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
            Left            =   5940
            TabIndex        =   28
            Top             =   255
            Width           =   705
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            Caption         =   "Condición alta"
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
            Left            =   5940
            TabIndex        =   27
            Top             =   1005
            Width           =   1125
         End
         Begin VB.Label lblFechaAlta 
            Caption         =   "Fecha alta"
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
            Top             =   960
            Width           =   1230
         End
         Begin VB.Label Label29 
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
            Left            =   120
            TabIndex        =   25
            Top             =   270
            Width           =   1155
         End
         Begin VB.Label Label43 
            Caption         =   "Médico egreso"
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
            TabIndex        =   24
            Top             =   1320
            Width           =   1335
         End
      End
      Begin VB.Frame Frame5 
         Height          =   615
         Left            =   -74865
         TabIndex        =   15
         Top             =   480
         Width           =   11475
         Begin VB.CommandButton btnMedicoRespNacimiento 
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
            Left            =   3210
            Picture         =   "AdmisionHospEgreso.frx":440B
            Style           =   1  'Graphical
            TabIndex        =   63
            Top             =   195
            Width           =   360
         End
         Begin VB.TextBox lblNombreMedicoNacimiento 
            Height          =   315
            Left            =   3585
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   195
            Width           =   7755
         End
         Begin VB.TextBox txtIdMedicoNacimiento 
            Height          =   315
            Left            =   2220
            MaxLength       =   10
            TabIndex        =   16
            Top             =   195
            Width           =   945
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Medico responsable"
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
            TabIndex        =   18
            Top             =   240
            Width           =   1590
         End
      End
      Begin VB.Frame Frame3 
         Height          =   705
         Left            =   -74850
         TabIndex        =   13
         Top             =   510
         Width           =   2985
         Begin VB.CheckBox chkSeRealizoNecropsia 
            Caption         =   "Se realizó necropsia?:"
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
            Left            =   450
            TabIndex        =   14
            Top             =   300
            Width           =   2445
         End
      End
      Begin UltraGrid.SSUltraGrid SSUltraGrid2 
         Height          =   4155
         Left            =   -74760
         TabIndex        =   41
         Top             =   510
         Width           =   11265
         _ExtentX        =   19870
         _ExtentY        =   7329
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108864
         Caption         =   "Lista de procedimientos"
      End
      Begin UltraGrid.SSUltraGrid SSUltraGrid1 
         Height          =   4155
         Left            =   -74790
         TabIndex        =   42
         Top             =   510
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   7329
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108864
         Caption         =   "Lista de examenes"
      End
      Begin SISGalenPlus.ucNacimientoDetalle ucNacimientoDetalle1 
         Height          =   3075
         Left            =   -74880
         TabIndex        =   43
         Top             =   1140
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   5424
      End
      Begin SISGalenPlus.ucDiagnosticoDetalle ucDiagnosticoNacimiento 
         Height          =   2325
         Left            =   -74910
         TabIndex        =   44
         Top             =   4260
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   3731
      End
      Begin SISGalenPlus.ucDiagnosticoDetalle ucDiagnosticoComplicaciones 
         Height          =   4035
         Left            =   -74865
         TabIndex        =   45
         Top             =   450
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   7117
      End
      Begin SISGalenPlus.ucDiagnosticoDetalle ucDiagnosticosEgreso 
         Height          =   3240
         Left            =   75
         TabIndex        =   7
         Top             =   4455
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   5715
      End
      Begin SISGalenPlus.ucDiagnosticoDetalle ucDiagnosticosMortalidad 
         Height          =   5205
         Left            =   -74880
         TabIndex        =   46
         Top             =   1320
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   9181
      End
      Begin SISGalenPlus.UcEpisodioClinico UcEpisodioClinico1 
         Height          =   930
         Left            =   5790
         TabIndex        =   47
         Top             =   2190
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   1640
      End
   End
   Begin SISGalenPlus.ucTransferenciasDetalle ucTransferenciasDetalle1 
      Height          =   2610
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   4604
   End
   Begin SISGalenPlus.ucDiagnosticoDetalle ucDiagnosticosIngreso 
      Height          =   2385
      Left            =   0
      TabIndex        =   49
      Top             =   2640
      Visible         =   0   'False
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   4207
   End
End
Attribute VB_Name = "AdmisionHospEgreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de la Alta Hospitalaria o Emergencia
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Teclado As New sighEntidades.Teclado
Dim mo_Formulario As New sighEntidades.Formulario
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim ml_idUsuario As Long
Dim mb_ExistenDatos As Boolean
Dim ml_TipoServicio As sghTipoServicio
Dim mo_Apariencia As New sighEntidades.GridInfragistic
Dim ml_EstadoCuenta As Long
'
Dim mo_sighProxies As New SIGHProxies.Procesos
Dim mo_ReglasSISgalenhos As New SIGHSis.ReglasSISgalenhos
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminServiciosGeograficos As New SIGHNegocios.ReglasServGeograf
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_AdminFacturacion As New ReglasFacturacion
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_AdminProgramacion As New SIGHNegocios.ReglasDeProgMedica
Dim ms_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
Dim mo_AdminHoteleria As New SIGHNegocios.ReglasHoteleria
Dim mo_Reniec As New ReniecGalenhos
Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
 
'
Dim mo_cmbIdTipoGravedad As New sighEntidades.ListaDespleglable
Dim mo_cmbIdEspecialidadMedico As New sighEntidades.ListaDespleglable
Dim mo_cmbIdServicio As New sighEntidades.ListaDespleglable
Dim mo_cmbIdDestinoAtencion As New sighEntidades.ListaDespleglable
Dim mo_cmbIdTipoReferenciaDestino As New sighEntidades.ListaDespleglable
Dim mo_cmbCondicionAlta As New sighEntidades.ListaDespleglable
Dim mo_cmbTipoAlta As New sighEntidades.ListaDespleglable
Dim mo_cmbServicioDestino As New sighEntidades.ListaDespleglable
'
Dim mo_DoUbicacionPaciente As New doPaciente
Dim mo_AtencionesEmergencia As New DOAtencionEmergencia
Dim mo_AtencionPadre As New DOAtencion
Dim mo_DoAtencionDatosAdicionales As New DoAtencionDatosAdicionales
Dim ldFechaEgresoMedicoAnterior As Date   'cuando se "modifique", generar "consumo por dias estancia"
Dim mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String

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
Dim ml_IdAtencionEmergencia As Long
Dim oRsEstancia As New Recordset
Dim mrs_OcupacionCamas As New ADODB.Recordset
Dim mo_OcupacionCamas As New Collection
Dim mrs_Interconsulta As New ADODB.Recordset
Dim ml_TipoAccionAdmision As sghTipoAccionEmergenciaYHospitalizacion
Dim ml_IdAtencionPadre As Long
'------------------------------------------------------------------------------------
'                               VARIABLE PARA LA ATENCION HOSPITALIZACION
'------------------------------------------------------------------------------------
Dim ml_IdAtencionHosp As Long
Dim mo_Diagnosticos As New Collection
Dim mo_Procedimientos As New Collection
Dim mo_Examenes As New Collection
Dim mo_Nacimientos As New Collection
Dim mo_NroServiciosQuePasoElPaciente As Long
Dim lcUltimoCodigoDeServicioTransferido As String

'------------------------------------------------------------------------------------
'                               VARIABLE PARA LA FILIACION
'------------------------------------------------------------------------------------
Dim ml_IdPaciente As Long
Dim mo_Pacientes  As New doPaciente
Dim ms_Autogenerado As String
Dim ml_TipoGeneracionHistoria As sghTipoNumeracionDeNroHistoria
Dim mo_Historia As New DOHistoriaClinica

'------------------------------------------------------------------------------------
'                               VARIABLE PARA LA CITA
'------------------------------------------------------------------------------------
Dim ml_IdMedico As Long
Dim ms_NombreMedico  As String
Dim mo_Especialidad As New DOEspecialidades
Dim mo_paciente As New doPaciente
Dim ml_IdPrestamo As Long
Dim ml_IdEspecialidad As Long

'------------------------------------------------------------------------------------
'                               PACIENTE NUEVO
'------------------------------------------------------------------------------------
Dim oRsFormaPago As New ADODB.Recordset
Dim oRsFuentesFinanciamiento As New ADODB.Recordset

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
Dim ml_IdServicioConCamaDisponible As Long
Dim lbPacienteNN As Boolean
Dim lnSecuenciaTransferencia As Long
Dim lnFocusCuandoCargeFrm As Long
Dim lbUsuarioConfirmaTransferencia As Boolean
Dim lbUsuarioConfirmaLlegada As Boolean
Dim lcCaptionTab2 As String
Dim lbUltimaTeclaPulsoENTER As Boolean
Dim lnIdPlanAnterior As Long
Dim lnIdTipoFinanciamientoAnterior As Long
Dim mo_lbCargaTablasUnaVez As Boolean
Dim mo_lbNuevoMovimiento As Boolean
Dim lnIdNacimientoSeleccionado As Long
Dim oDoSunasaPacientesHistoricos As New DoSunasaPacientesHistoricos
Dim mo_DOAtencionesCE As New DOAtencionesCE
Dim mb_EsObservacionEmergencia As Boolean
Dim lbBuscaDNIenReniec As Boolean
Dim lcElServicioUsaGalenHos As String
Dim ldFechaActualServidor As Date
Const lbCargaAlaVezCitaPacienteAtencionDA As Boolean = True
Dim ml_ldFechaIngreso As Date, ml_lcHoraIngreso As String
Dim ml_IdFuenteFinanciamiento As Long, ml_IdViasAdmision As Long
Dim mc_FuaVersionFormato As String
Dim lcHistoriaYpaciente As String
Dim ml_TipoFinanciamiento As String
Dim lnIdTipoGravedadEgreso As Long
Dim lbTieneLicenciaParaMensajeAcelulares As Boolean
Dim lbSeEnviaMensajeCelularAreaConv As Boolean


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
   'mgaray09
   setListItemAControlDiagnosticos mo_lnIdTablaLISTBARITEMS
End Property
Property Let IdServicioConCamaDisponible(lValue As Long)
    ml_IdServicioConCamaDisponible = lValue
End Property
Property Get IdServicioConCamaDisponible() As Long
    IdServicioConCamaDisponible = ml_IdServicioConCamaDisponible
End Property


Property Let idMedico(lValue As Long)
   ml_IdMedico = lValue
End Property
Property Get idMedico() As Long
   idMedico = ml_IdMedico
End Property
Property Let IdPrestamo(lValue As Long)
   ml_IdPrestamo = lValue
End Property
Property Get IdPrestamo() As Long
   IdPrestamo = ml_IdPrestamo
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
Property Let IdAtencionHosp(lValue As Long)
   ml_IdAtencionHosp = lValue
End Property
Property Get IdAtencionHosp() As Long
   IdAtencionHosp = ml_IdAtencionHosp
End Property
Property Let IdAtencionEmergencia(lValue As Long)
   ml_IdAtencionEmergencia = lValue
End Property
Property Get IdAtencionEmergencia() As Long
   IdAtencionEmergencia = ml_IdAtencionEmergencia
End Property
Property Let idPaciente(lValue As Long)
   ml_IdPaciente = lValue
End Property
Property Get idPaciente() As Long
   idPaciente = ml_IdPaciente
End Property
Property Let TipoServicio(sValue As sghTipoServicio)
   ml_TipoServicio = sValue
End Property
Property Get TipoServicio() As sghTipoServicio
   TipoServicio = ml_TipoServicio
End Property
Property Let TipoGeneracionHistoria(lValue As Long)
   ml_TipoGeneracionHistoria = lValue
End Property
Property Get TipoGeneracionHistoria() As Long
   TipoGeneracionHistoria = ml_TipoGeneracionHistoria
End Property
Property Let IdEspecialidad(lValue As Long)
   ml_IdEspecialidad = lValue
End Property
Property Get IdEspecialidad() As Long
   IdEspecialidad = ml_IdEspecialidad
End Property
Property Let TipoAccionDeAdmision(lValue As sghTipoAccionEmergenciaYHospitalizacion)
    ml_TipoAccionAdmision = lValue
End Property
Property Get TipoAccionDeAdmision() As sghTipoAccionEmergenciaYHospitalizacion
    TipoAccionDeAdmision = ml_TipoAccionAdmision
End Property

Property Let IdAtencionPadre(lValue As Long)
   ml_IdAtencionPadre = lValue
End Property
Property Get IdAtencionPadre() As Long
   IdAtencionPadre = ml_IdAtencionPadre
End Property



Sub CargarComboBoxes()
Dim sSQL As String
Dim sMensaje As String

       mo_cmbIdDestinoAtencion.BoundColumn = "IdDestinoAtencion"
       mo_cmbIdDestinoAtencion.ListField = "DescripcionLarga"
       Select Case ml_TipoServicio
       Case 2
            Set mo_cmbIdDestinoAtencion.RowSource = mo_AdminAdmision.TiposDestinoAtencionSeleccionarDestinosDeConsultorioEmergencia
       Case 3
            Set mo_cmbIdDestinoAtencion.RowSource = mo_AdminAdmision.TiposDestinoAtencionSeleccionarDestinosDeHospitalizacion(sghSoloPacHospitalizados)
       End Select

       
       mo_cmbIdTipoReferenciaDestino.BoundColumn = "IdTipoReferencia"
       mo_cmbIdTipoReferenciaDestino.ListField = "DescripcionLarga"
       Set mo_cmbIdTipoReferenciaDestino.RowSource = mo_AdminServiciosComunes.TiposReferenciaSeleccionarTodos
       sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError
       
        
       mo_cmbCondicionAlta.BoundColumn = "IdCondicionAlta"
       mo_cmbCondicionAlta.ListField = "DescripcionLarga"
       Set mo_cmbCondicionAlta.RowSource = mo_AdminServiciosComunes.TiposCondicionAltaSeleccionarTodos
       sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError
        
       mo_cmbTipoAlta.BoundColumn = "IdTipoAlta"
       mo_cmbTipoAlta.ListField = "DescripcionLarga"
       Set mo_cmbTipoAlta.RowSource = mo_AdminServiciosComunes.TiposAltaSeleccionarTodos
       sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError
       
       If ml_TipoServicio = sghEmergenciaConsultorios Then
          mo_cmbServicioDestino.BoundColumn = "IdServicio"
          mo_cmbServicioDestino.ListField = "DservicioHosp"
          Set mo_cmbServicioDestino.RowSource = mo_AdminAdmision.DevuelveServiciosDelHospital("(3)", "", sghFiltraSoloActivos, sghPorDescServicio)
       End If
       
       If sMensaje <> "" Then
           MsgBox sMensaje, vbInformation, Me.Caption
       End If
     

        
        Me.ucDiagnosticosEgreso.TipoDiagnostico = sghHospitalizacionEgreso
        Me.ucDiagnosticosEgreso.IdListBarItem = mo_lnIdTablaLISTBARITEMS
        Me.ucDiagnosticoNacimiento.TipoDiagnostico = sghHospitalizacionNacimiento
        Me.ucDiagnosticoNacimiento.IdListBarItem = mo_lnIdTablaLISTBARITEMS
        Me.ucDiagnosticoComplicaciones.TipoDiagnostico = sghHospitalizacionComplicaciones
        Me.ucDiagnosticoComplicaciones.IdListBarItem = mo_lnIdTablaLISTBARITEMS
        Me.ucDiagnosticosMortalidad.TipoDiagnostico = sghHospitalizacionMortalidad
        Me.ucDiagnosticosMortalidad.IdListBarItem = mo_lnIdTablaLISTBARITEMS

        Me.ucDiagnosticosEgreso.ConfigurarComboBoxes
        Me.ucDiagnosticoNacimiento.ConfigurarComboBoxes
        Me.ucDiagnosticosMortalidad.ConfigurarComboBoxes
        Me.ucDiagnosticoComplicaciones.ConfigurarComboBoxes
        Me.ucNacimientoDetalle1.ConfigurarComboBoxes
        '
        Set cmbServicioReferenciaD.ListSource = mo_AdminServiciosComunes.SuSalud_upsSeleccionarTodos   'debb-21/06/2016
        '
End Sub





Private Sub btnBuscaHistoricos_Click()
    Dim oBuscaHistoricos As New AdmisionCEhistorico
    oBuscaHistoricos.MuestraTab = 2
    oBuscaHistoricos.Paciente = lcHistoriaYpaciente
    oBuscaHistoricos.idPaciente = ml_IdPaciente
    oBuscaHistoricos.idTipoSexo = mo_paciente.idTipoSexo
    oBuscaHistoricos.NroHistoriaClinica = Val(Mid(lcHistoriaYpaciente, 2, InStr(lcHistoriaYpaciente, ")") - 2))
    oBuscaHistoricos.Show 1
    Set oBuscaHistoricos = Nothing
    
End Sub

Private Sub btnBuscarEstablecimientoDestino_Click()
    If cmbIdTipoReferenciaDestino.Text <> "" Then
       CompletarDatosDeEstablecimiento txtIdEstablecimientoDestino, lblNombreDestinoReferencia, mo_cmbIdTipoReferenciaDestino.BoundText
    End If
End Sub

Private Sub btnBuscarMedicosEgreso_Click()
    CompletarDatosDeMedicoEgreso txtIdMedicoEgreso, lblNombreMedicoEgreso, Val(Me.lblNombreServicioEgreso.Tag), "", CDate(Me.txtFechaEgreso.Text), Me.txtHoraEgreso.Text, ml_TipoServicio
End Sub


Private Sub btnImprimeFichaSIS_Click()
    If Me.txtFechaEgreso.Text <> sighEntidades.FECHA_VACIA_DMY Then
        If mo_Atenciones.idTipoServicio = sghTipoServicio.sghEmergenciaConsultorios And _
           mo_Atenciones.IdDestinoAtencion = 21 And wxParametro512 = "S" Then
           MsgBox "Se GRABO CORRECTAMENTE, pero el FUA se emitirá en HOSPITALIZACION -> ALTA MEDICA", vbInformation, Me.Caption
        Else
            Dim ml_FuaTipoAnexo2015 As Integer
            Dim oFua As New SIGHSis.clFUA
            oFua.idCuentaAtencion = ml_idCuentaAtencion
            oFua.lcNombrePc = mo_lcNombrePc
            oFua.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
            oFua.idUsuario = ml_idUsuario
            oFua.EsAltaMedica = True
            oFua.IdServicio = CLng(Me.txtIdServicioEgreso.Tag)
            oFua.MostrarFormulario
            Set oFua = Nothing
        End If
    Else
        MsgBox "Aún no se GRABO el ALTA MEDICA", vbInformation, Me.Caption
    End If
End Sub

Private Sub btnMedicoRespNacimiento_Click()
    'Buscará Médicos por Especialidad y no los Programados, porque no se sabe la Fecha/hora Nacimiento
    CompletarDatosDeMedico Me.txtIdMedicoNacimiento, Me.lblNombreMedicoNacimiento, 0, "", ml_ldFechaIngreso, ml_lcHoraIngreso, 0
End Sub

Private Sub cmbCondicionAlta_Click()
        Me.TabEgresos.TabVisible(3) = Val(mo_cmbCondicionAlta.BoundText) = 4
End Sub
































Private Sub cmbServicioReferenciaD_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbServicioReferenciaD
End Sub

Private Sub cmdBuscaCamaEgreso_Click()
Dim oBusqueda As New CamasBusqueda
Dim oDOCama As New DOCama
Dim oConexion As New Connection
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    oBusqueda.idTipoServicio = ml_TipoServicio
    oBusqueda.IdServicio = Val(txtIdServicioEgreso.Tag)
    oBusqueda.Show 1
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOCama = mo_AdminHoteleria.CamasSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oConexion)
        If Not oDOCama Is Nothing Then
            If oDOCama.idPaciente = mo_Atenciones.idPaciente Or oDOCama.idPaciente = 0 Then
                Me.txtNroCamaEgreso.Text = oDOCama.Codigo
                Me.txtNroCamaEgreso.Tag = oDOCama.idCama
            Else
                MsgBox "La cama seleccionada no puede usarla", vbInformation, Me.Caption
            End If
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oBusqueda = Nothing
End Sub











Private Sub cmdImpresionReferencias_Click()
    Dim sCodigoDestino As String
    If cmbIdDestinoAtencion.Text = "" Then
       Exit Sub
    End If
    
    sCodigoDestino = Trim(Split(cmbIdDestinoAtencion.Text, " = ")(0))
    If sCodigoDestino = "R" Or sCodigoDestino = "C" Then
       If ValidarDatosObligatorios() Then
            CargaDatosAlObjetosDeDatos False
            If ValidarReglas() Then
                Dim oReferencias As New SIGHReportes.clReferencias
                If sCodigoDestino = "C" Then
                   oReferencias.CreaReporteContrarefencias lblNombreDestinoReferencia.Text, mo_paciente, mo_Atenciones, _
                                                            mo_DoAtencionDatosAdicionales, IIf(chkEnPDF.Value = 1, True, False)
                Else
                   oReferencias.CrearReporteReferencias lblNombreDestinoReferencia.Text, mo_paciente, mo_Atenciones, _
                                                            mo_DoAtencionDatosAdicionales, IIf(chkEnPDF.Value = 1, True, False)
                End If
                Set oReferencias = Nothing
            End If
       End If
     End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub



















Private Sub lblNombreMedicoEgreso_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, lblNombreMedicoEgreso
   AdministrarKeyPreview KeyCode

End Sub

Private Sub lblNombreMedicoEgreso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lbUltimaTeclaPulsoENTER = True
    Else
        lbUltimaTeclaPulsoENTER = False
    End If

End Sub

Private Sub lblNombreMedicoEgreso_LostFocus()
    If Me.txtFechaEgreso.Text <> sighEntidades.FECHA_VACIA_DMY Then
    If lbUltimaTeclaPulsoENTER = True Then
       lbUltimaTeclaPulsoENTER = False
       CompletarDatosDeMedicoEgreso txtIdMedicoEgreso, lblNombreMedicoEgreso, Val(Me.lblNombreServicioEgreso.Tag), lblNombreMedicoEgreso.Text, CDate(Me.txtFechaEgreso.Text), Me.txtHoraEgreso.Text, ml_TipoServicio
    End If
    mo_Formulario.MarcarComoVacio lblNombreMedicoEgreso
    End If
End Sub






















Private Sub txtFextension_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFextension
End Sub



Private Sub txtFtramite_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFtramite
End Sub

Private Sub txtHoraEgreso_Change()
   If mi_Opcion = sghAgregar Then
      lblNombreMedicoEgreso.Text = ""
      txtIdMedicoEgreso.Text = ""
   End If
End Sub

Private Sub cmbCondicionAlta_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbCondicionAlta
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbCondicionAlta_LostFocus()
    If cmbCondicionAlta.Text <> "" Then
       mo_cmbCondicionAlta.BoundText = Val(Split(cmbCondicionAlta.Text, " = ")(0))
    End If
    mo_Formulario.MarcarComoVacio cmbCondicionAlta
    If txtHoraEgreso <> sighEntidades.HORA_VACIA_HM Then
        On Error Resume Next
        If mo_cmbCondicionAlta.BoundText <> "4" Then
           ucDiagnosticosEgreso.SetFocus
        Else
           'fallecido
           ucDiagnosticosEgreso.LimpiarDatos
           TabEgresos.Tab = 3
           ucDiagnosticosMortalidad.SetFocus
        End If
    End If
End Sub

Private Sub cmbCondicionAlta_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub cmbIdDestinoAtencion_Click()
Dim sCodigoDestino As String
    If cmbIdDestinoAtencion.Text <> "" Then
       sCodigoDestino = Trim(Split(cmbIdDestinoAtencion.Text, " = ")(0))
    Else
       sCodigoDestino = " "
    End If
    If sCodigoDestino <> "R" And sCodigoDestino <> "C" Then
        mo_cmbIdTipoReferenciaDestino.BoundText = ""
        Me.txtIdEstablecimientoDestino.Tag = ""
        Me.txtIdEstablecimientoDestino = ""
        Me.lblNombreDestinoReferencia = ""
        txtReferenciaD.Text = ""
         'debb-21/06/2016 (inicio)
        cmbServicioReferenciaD.Text = ""
        txtFextension.Text = sighEntidades.FECHA_VACIA_DMY
        txtFtramite.Text = sighEntidades.FECHA_VACIA_DMY
         'debb-21/06/2016 (fin)

    Else
        
        mo_cmbIdTipoReferenciaDestino.BoundText = "1"
        txtReferenciaD.Text = mo_AdminServiciosComunes.CalculaNUMEROREFERENCIA(IIf(sCodigoDestino = "C", True, False))
        mo_cmbIdTipoReferenciaDestino.BoundText = "1"
        
        txtIdEstablecimientoDestino.Tag = ""
        txtIdEstablecimientoDestino.Text = ""
        lblNombreDestinoReferencia = ""
        If mo_DoAtencionDatosAdicionales.IdEstablecimientoOrigen > 0 And sCodigoDestino = "C" Then
            Dim oDoEstablecimiento As New DOEstablecimiento
            Set oDoEstablecimiento = mo_AdminServiciosComunes.EstablecimientosSeleccionarPorId(mo_DoAtencionDatosAdicionales.IdEstablecimientoOrigen)
            If Not oDoEstablecimiento Is Nothing Then
                txtIdEstablecimientoDestino.Tag = oDoEstablecimiento.IdEstablecimiento
                txtIdEstablecimientoDestino.Text = oDoEstablecimiento.Codigo
                lblNombreDestinoReferencia = oDoEstablecimiento.nombre
            End If
            Set oDoEstablecimiento = Nothing
        End If
    End If
    HabilitarFrameDestino False
    '
    If ml_TipoServicio = sghEmergenciaConsultorios Then
        If sCodigoDestino = "H" Then
           mo_Formulario.HabilitarDeshabilitar cmbServicioDestino, True
        Else
           mo_cmbServicioDestino.BoundText = ""
           mo_Formulario.HabilitarDeshabilitar cmbServicioDestino, False
        End If
    End If
    '
    cmbTipoAlta.Text = ""
    cmbCondicionAlta.Text = ""
'    Select Case sCodigoDestino
'    Case "R"
'        HabilitarFrameDestino True
'        Me.fraDatosReferenciaDestino = "Refer.Destino"
'        Me.lblIdTipoReferenciaDestino = "Tipo Referencia"
'        Me.lblIdEstablecimientoDestino = "Estab.Refer"
'        mo_cmbTipoAlta.BoundText = "4"
'    Case "C"
'        HabilitarFrameDestino True
'        Me.fraDatosReferenciaDestino = "Contraref.Destino"
'        Me.lblIdTipoReferenciaDestino = "Tipo Contraref"
'        Me.lblIdEstablecimientoDestino = "Estab. Contraref"
'        mo_cmbTipoAlta.BoundText = "4"
'    Case "M", "U"
'        mo_cmbCondicionAlta.BoundText = "4"
'    Case "O"
'        lblFechaAlta.Caption = "Fecha traslado"
'        lblFechaEgresoAdm.Visible = False
'        txtFechaEgresoAdm.Visible = False
'        txtHoraEgresoAdm.Visible = False
'        mo_cmbTipoAlta.BoundText = "5"
'    Case "H"
'        lblFechaAlta.Caption = "Fecha traslado"
'        lblFechaEgresoAdm.Visible = False
'        txtFechaEgresoAdm.Visible = False
'        txtHoraEgresoAdm.Visible = False
'        mo_cmbTipoAlta.BoundText = "6"
'    Case "D"
'        mo_cmbTipoAlta.BoundText = "1"
'    Case "F"
'        mo_cmbTipoAlta.BoundText = "3"
'    Case Else
'        lblFechaAlta.Caption = "Fecha alta"
'        lblFechaEgresoAdm.Visible = True
'        txtFechaEgresoAdm.Visible = True
'        txtHoraEgresoAdm.Visible = True
'    End Select

'Actualizado 16102014
    Select Case sCodigoDestino
    Case "D"
        lblFechaAlta.Caption = "Fecha Alta" ' A.Yañez 14/10/2014
        txtFechaEgresoAdm.Visible = False 'A.Yañez 14/10/2014
        txtHoraEgresoAdm.Visible = False 'A.Yañez 14/10/2014
    Case "R"
        HabilitarFrameDestino True
        Me.fraDatosReferenciaDestino = "Refer.Destino"
        Me.lblIdTipoReferenciaDestino = "Tipo Referencia"
        Me.lblIdEstablecimientoDestino = "Estab.Refer"
        mo_cmbTipoAlta.BoundText = "4"
        txtFechaEgresoAdm.Visible = False 'A.Yañez 14/10/2014
        txtHoraEgresoAdm.Visible = False 'A.Yañez 14/10/2014
    Case "C"
        HabilitarFrameDestino True
        Me.fraDatosReferenciaDestino = "Contraref.Destino"
        Me.lblIdTipoReferenciaDestino = "Tipo Contraref"
        Me.lblIdEstablecimientoDestino = "Estab. Contraref"
        mo_cmbTipoAlta.BoundText = "4"
        txtFechaEgresoAdm.Visible = False 'A.Yañez 14/10/2014
        txtHoraEgresoAdm.Visible = False 'A.Yañez 14/10/2014
    Case "M", "U"
         mo_cmbTipoAlta.BoundText = "7"      'debb-15/06/2016
        mo_cmbCondicionAlta.BoundText = "4"
        txtFechaEgresoAdm.Visible = False 'A.Yañez 14/10/2014
        txtHoraEgresoAdm.Visible = False 'A.Yañez 14/10/2014
        
    Case "O"
        lblFechaAlta.Caption = "Fecha Traslado"
        lblFechaEgresoAdm.Visible = False
        txtFechaEgresoAdm.Visible = False 'A.Yañez 14/10/2014
        txtHoraEgresoAdm.Visible = False 'A.Yañez 14/10/2014
        mo_cmbTipoAlta.BoundText = "5"
    Case "H"
        lblFechaAlta.Caption = "Fecha Traslado"
        lblFechaEgresoAdm.Visible = False
        txtFechaEgresoAdm.Visible = False 'A.Yañez 14/10/2014
        txtHoraEgresoAdm.Visible = False 'A.Yañez 14/10/2014
        mo_cmbTipoAlta.BoundText = "6"
    Case "D"
        mo_cmbTipoAlta.BoundText = "1"
        txtFechaEgresoAdm.Visible = False 'A.Yañez 14/10/2014
        txtHoraEgresoAdm.Visible = False 'A.Yañez 14/10/2014
    Case "F"
        mo_cmbTipoAlta.BoundText = "3"
        txtFechaEgresoAdm.Visible = False 'A.Yañez 14/10/2014
        txtHoraEgresoAdm.Visible = False 'A.Yañez 14/10/2014
    Case Else
        lblFechaAlta.Caption = "Fecha Alta"
        lblFechaEgresoAdm.Visible = True
        txtFechaEgresoAdm.Visible = True
        txtHoraEgresoAdm.Visible = True
    End Select


End Sub
Sub HabilitarFrameDestino(bValue As Boolean)
        mo_Formulario.HabilitarDeshabilitar fraDatosReferenciaDestino, bValue
        mo_Formulario.HabilitarDeshabilitar fraDatosReferenciaDestino, bValue
        mo_Formulario.HabilitarDeshabilitar lblIdTipoReferenciaDestino, bValue
        mo_Formulario.HabilitarDeshabilitar cmbIdTipoReferenciaDestino, bValue
        mo_Formulario.HabilitarDeshabilitar lblIdEstablecimientoDestino, bValue
        mo_Formulario.HabilitarDeshabilitar Me.txtReferenciaD, bValue
        'debb-21/06/2016 (inicio)
        mo_Formulario.HabilitarDeshabilitar cmbServicioReferenciaD, bValue
        mo_Formulario.HabilitarDeshabilitar txtFextension, bValue
        mo_Formulario.HabilitarDeshabilitar txtFtramite, bValue
        mo_Formulario.HabilitarDeshabilitar Me.lblFtramite0, bValue
        mo_Formulario.HabilitarDeshabilitar Me.lblFextension0, bValue
        mo_Formulario.HabilitarDeshabilitar lblServicioReferencia0, bValue
        mo_Formulario.HabilitarDeshabilitar lblNreferencia0, bValue
        'debb-21/06/2016 (fin)
End Sub

Private Sub cmbIdDestinoAtencion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdDestinoAtencion
   AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdDestinoAtencion_LostFocus()
Dim oDOTipoDestinoAtencion As New DOTipoDestinoAtencion
Dim lcValorP As String
Dim lnIdUsuarioMedico As Long
Dim oDoMedico As New DOMedico
Dim oMedicosEspecialidad As New Collection
Dim oDOEmpleado As New dOEmpleado
Dim oMedicoEspecialidad As New MedicosEspecialidad
Dim oDoMedicoEspecialidad As New DOMedicoEspecialidad
Dim oRsTmp As New Recordset
Dim rsMedicoEspecialidad As New Recordset
Dim lbMismaEspecialidad As Boolean
Dim oConexion As New Connection
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighEntidades.CadenaConexion
'Dim oDOTipoDestinoAtencion As New DOTipoDestinoAtencion
'Dim lcValorP As String
    If cmbIdDestinoAtencion.Text <> "" Then
         Set oDOTipoDestinoAtencion = mo_AdminAdmision.TiposDestinoAtencionSeleccionarPorCodigo(Trim(Split(cmbIdDestinoAtencion.Text, " = ")(0)), ml_TipoServicio)
         If oDOTipoDestinoAtencion.IdDestinoAtencion <> 0 Then
             mo_cmbIdDestinoAtencion.BoundText = oDOTipoDestinoAtencion.IdDestinoAtencion
        End If
    End If
    If cmbIdDestinoAtencion.Text <> "" Then
      If txtFechaEgreso.Text = sighEntidades.FECHA_VACIA_DMY Or Not (IsDate(txtFechaEgreso.Text)) Then txtFechaEgreso.Text = lcBuscaParametro.RetornaFechaServidorSQL
      
'      If txtHoraEgreso.Text = sighentidades.HORA_VACIA_HM Or Not (IsDate(txtHoraEgreso.Text)) Then txtHoraEgreso.Text = lcBuscaParametro.RetornaHoraServidorSQL
      'Modificado Yamill 08092014
      If txtHoraEgreso.Text = sighEntidades.HORA_VACIA_HM Or Not (IsDate(txtHoraEgreso.Text)) Then txtHoraEgreso.Text = lcBuscaParametro.RetornaHoraServidorSQLserverFormatoGalenhos
      
      '
      
            'Yamill Palomino
      
      'si el usuario de login es medico
        'si la especialidad del usuario medico es igual a la especialidad dond esta siendo atendido
            'cargamos los datos del usuario al medico de egreso
        'sino
            'limpiamos
        'sino
      'sino
        
       Set oRsTmp = mo_ReglasDeProgMedica.MedicosXidEmpleado(ml_idUsuario, oConexion)
       If oRsTmp.RecordCount > 0 Then
            lnIdUsuarioMedico = oRsTmp.Fields!idMedico
            If mo_AdminProgramacion.MedicosSeleccionarPorId(lnIdUsuarioMedico, oDoMedico, oDOEmpleado, oMedicosEspecialidad, oConexion) Then
                Set oMedicoEspecialidad.Conexion = oConexion
                Set rsMedicoEspecialidad = oMedicoEspecialidad.SeleccionarPorMedico(lnIdUsuarioMedico)
                lbMismaEspecialidad = False
                If rsMedicoEspecialidad.RecordCount > 0 Then
                   rsMedicoEspecialidad.MoveFirst
                   If Me.lblNombreServicioEgreso.Tag <> "" Then
                        rsMedicoEspecialidad.Find "IdEspecialidad=" & Val(Me.lblNombreServicioEgreso.Tag)
                   Else
                        rsMedicoEspecialidad.Find "IdEspecialidad=" & ml_IdEspecialidad
                   End If
                   If Not rsMedicoEspecialidad.EOF Then
                      lbMismaEspecialidad = True
                   End If
                End If
                If lbMismaEspecialidad Then
                    txtIdMedicoEgreso.Text = oDOEmpleado.CodigoPlanilla
                    txtIdMedicoEgreso.Tag = oDoMedico.idMedico
                    lblNombreMedicoEgreso = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
                Else
                   ' Me.lblNombreMedicoEgreso = ""
                   ' txtIdMedicoEgreso.Text = ""
                End If
            End If
        Else
           ' Me.lblNombreMedicoEgreso = ""
           ' txtIdMedicoEgreso.Text = ""
        End If
        
        '
      
      If txtHoraEgreso.Text = sighEntidades.HORA_VACIA_HM Or Not (IsDate(txtHoraEgreso.Text)) Then txtHoraEgreso.Text = lcBuscaParametro.RetornaHoraServidorSQLserverFormatoGalenhos
      '
      If cmbTipoAlta.Text = "" Then                                     'debb-06-03-2012
         lcValorP = wxParametro290
         If Val(lcValorP) > 0 Then
            mo_cmbTipoAlta.BoundText = lcValorP
         End If
      End If
      '
      If cmbCondicionAlta.Text = "" Then                                'debb-06-03-2012
         lcValorP = wxParametro291
         If Val(lcValorP) > 0 Then
            mo_cmbCondicionAlta.BoundText = lcValorP
         End If
      End If
      '
      lcValorP = wxParametro292              'debb-06-03-2012
      If Val(lcValorP) > 0 Then
         ucDiagnosticosEgreso.TipoDxDefault lcValorP
      End If
      '
    End If
    mo_Formulario.MarcarComoVacio cmbIdDestinoAtencion
    Set oDOTipoDestinoAtencion = Nothing
    Set oMedicosEspecialidad = Nothing
    Set oDOTipoDestinoAtencion = Nothing
    Set oDOEmpleado = Nothing
    Set oDoMedico = Nothing
    Set oMedicoEspecialidad = Nothing
    Set oDoMedicoEspecialidad = Nothing
    Set oRsTmp = Nothing
    Set rsMedicoEspecialidad = Nothing
    Set oConexion = Nothing

End Sub

Private Sub cmbIdDestinoAtencion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
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

Private Sub cmbTipoAlta_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbTipoAlta
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbTipoAlta_LostFocus()
   If cmbTipoAlta.Text <> "" Then
       mo_cmbTipoAlta.BoundText = Val(Split(cmbTipoAlta.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbTipoAlta
End Sub

Private Sub cmbTipoAlta_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub



Private Sub Form_Initialize()
    
    Set mo_cmbIdTipoGravedad.MiComboBox = cmbIdTipoGravedad
    Set mo_cmbIdDestinoAtencion.MiComboBox = cmbIdDestinoAtencion
    Set mo_cmbIdTipoReferenciaDestino.MiComboBox = cmbIdTipoReferenciaDestino
    Set mo_cmbCondicionAlta.MiComboBox = cmbCondicionAlta
    Set mo_cmbTipoAlta.MiComboBox = cmbTipoAlta
    Set mo_cmbServicioDestino.MiComboBox = cmbServicioDestino

End Sub





Private Sub txtFechaEgreso_Change()

'
End Sub

Private Sub txtFechaEgreso_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFechaEgreso
   AdministrarKeyPreview KeyCode
End Sub


Private Sub txtFechaEgreso_LostFocus()
       If txtFechaEgreso <> sighEntidades.FECHA_VACIA_DMY Then
            If Not EsFecha(txtFechaEgreso, "DD/MM/AAAA") Then
                MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
                 txtFechaEgreso = sighEntidades.FECHA_VACIA_DMY
            End If
        End If
        
        mo_Formulario.MarcarComoVacio txtFechaEgreso
End Sub

Private Sub txtFechaEgreso_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaFecha(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtFechaEgresoAdm_KeyDown(KeyCode As Integer, Shift As Integer)
        
        mo_Teclado.RealizarNavegacion KeyCode, txtFechaEgresoAdm
        AdministrarKeyPreview KeyCode
End Sub


Private Sub txtFechaEgresoAdm_LostFocus()
       If txtFechaEgresoAdm <> sighEntidades.FECHA_VACIA_DMY Then
            If Not EsFecha(txtFechaEgresoAdm, "DD/MM/AAAA") Then
                MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
                 txtFechaEgresoAdm = sighEntidades.FECHA_VACIA_DMY
            End If
        End If
           
   mo_Formulario.MarcarComoVacio txtFechaEgresoAdm
End Sub

Private Sub txtFechaEgresoAdm_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaFecha(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtIdDiagnostico_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
        If Not mo_Teclado.CodigoAsciiEsCIE10(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub



Private Sub txtHoraEgreso_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtHoraEgreso
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtHoraEgreso_LostFocus()
        If txtHoraEgreso <> sighEntidades.HORA_VACIA_HM Then
            If Not sighEntidades.ValidaHora(txtHoraEgreso) Then
                MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
                txtHoraEgreso = sighEntidades.HORA_VACIA_HM
            End If
        End If
        mo_Formulario.MarcarComoVacio txtHoraEgreso
        Me.lblNombreMedicoEgreso.SetFocus
End Sub

Private Sub txtHoraEgreso_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaHora(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtHoraEgresoAdm_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtHoraEgresoAdm
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtHoraEgresoAdm_LostFocus()
        
        If txtHoraEgresoAdm <> sighEntidades.HORA_VACIA_HM Then
            If Not sighEntidades.ValidaHora(txtHoraEgresoAdm) Then
                MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
                 txtHoraEgresoAdm.Text = sighEntidades.HORA_VACIA_HM
            End If
        End If
        
        mo_Formulario.MarcarComoVacio txtHoraEgresoAdm
End Sub

Private Sub txtHoraEgresoAdm_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaHora(KeyAscii) Then
           KeyAscii = 0
       End If
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
    CompletarDatosDelEstablecimientoEnElLostFocus txtIdEstablecimientoDestino, lblNombreDestinoReferencia, Val(mo_cmbIdTipoReferenciaDestino.BoundText)
    mo_Formulario.MarcarComoVacio txtIdEstablecimientoDestino
End Sub

Private Sub txtIdEstablecimientoDestino_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub






























Private Sub txtIdMedicoEgreso_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdMedicoEgreso
    If KeyCode = vbKeyF1 Then
        btnBuscarMedicosEgreso_Click
    End If
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtIdMedicoEgreso_LostFocus()
    CompletarDatosDeMedicoEnElLostFocus txtIdMedicoEgreso, lblNombreMedicoEgreso
    mo_Formulario.MarcarComoVacio txtIdMedicoEgreso
End Sub

Private Sub txtIdMedicoEgreso_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub







Private Sub txtIdMedicoNacimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdMedicoNacimiento
    If KeyCode = vbKeyF1 Then
        btnMedicoRespNacimiento_Click
    End If
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtIdMedicoNacimiento_LostFocus()
    CompletarDatosDeMedicoEnElLostFocus txtIdMedicoNacimiento, lblNombreMedicoEgreso
    mo_Formulario.MarcarComoVacio txtIdMedicoNacimiento
End Sub

Private Sub txtIdMedicoNacimiento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtIdServicioEgreso_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdServicioEgreso
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdServicioEgreso_LostFocus()
    CompletarDatosDeServicioEnElLostFocus txtIdServicioEgreso, lblNombreServicioEgreso
    mo_Formulario.MarcarComoVacio txtIdServicioEgreso
End Sub

Private Sub txtIdServicioEgreso_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub















Sub CargarDatosAlFormulario()
    Dim lnIdCamaSeleccionada As Long
    mo_Formulario.HabilitarDeshabilitar Me.lblNombreMedicoEgreso, True
    mo_Formulario.HabilitarDeshabilitar txtIdMedicoEgreso, False
    mo_Formulario.HabilitarDeshabilitar Me.txtIdServicioEgreso, False
    mo_Formulario.HabilitarDeshabilitar Me.lblNombreServicioEgreso, False
    mo_Formulario.HabilitarDeshabilitar Me.lblNombreMedicoNacimiento, False
    mo_Formulario.HabilitarDeshabilitar Me.txtNroCamaEgreso, False
    mo_Formulario.HabilitarDeshabilitar Me.txtIdEstablecimientoDestino, False
    mo_Formulario.HabilitarDeshabilitar lblNombreDestinoReferencia, False   'debb-05/04/2011
    
    
    If ml_TipoServicio <> sghHospitalizacion Then Me.cmdBuscaCamaEgreso.Visible = False
    
    Me.ucTransferenciasDetalle1.TipoServicio = ml_TipoServicio
    
    Me.ucDiagnosticosEgreso.TituloFrame = "Diagnósticos de Egreso     (F1=Todos Dx)"
    Me.ucDiagnosticoNacimiento.TituloFrame = "Diagnósticos de muerte fetal     (F1=Todos Dx)"
    Me.ucDiagnosticoComplicaciones.TituloFrame = "Complicaciones     (F1=Todos Dx)"
    Me.ucDiagnosticosMortalidad.TituloFrame = "Diagnósticos de mortalidad     (F1=Todos Dx)"
    
    lcUltimoCodigoDeServicioTransferido = ""
        
    Select Case ml_TipoAccionAdmision
    Case sghAdmisionNormal  'Si el una admisión normal de hospitalizacion o de emergencia
        Select Case mi_Opcion
            Case sghModificar
                CargarDatosAlosControles
        End Select
        
    Case sghEnviarAObservacion
        mi_Opcion = sghAgregar
        ValoresPorDefecto
        CargarDatosParaEnviarAObservacion
        
    Case sghTrasladarAHospitalizacion
        mi_Opcion = sghAgregar
        ValoresPorDefecto
        CargarDatosParaEnviarAHospitalizacion
        
    Case sghDarDeAlta
        mi_Opcion = sghModificar
        CargarDatosAlosControles
        
    Case sghIngresarUnAlojamientoConjunto
        mi_Opcion = sghAgregar
        ValoresPorDefecto
        
        
    Case sghTransferencias
        mi_Opcion = sghModificar
        CargarDatosAlosControles
        
    End Select
    
     Select Case mi_Opcion
     Case sghModificar
     
     Case sghConsultar
    
    Case sghEliminar
    
    End Select
    
End Sub
Sub CargarDatosParaEnviarAObservacion()
        
        CargarDatosDeLasAtencionesPadres
        

        
        lcBuscaParametro.RetornaHoraServidorSQL1
        ml_ldFechaIngreso = lcBuscaParametro.RetornaFechaServidorSQL 'Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY)
        ml_lcHoraIngreso = lcBuscaParametro.RetornaHoraServidorSQLserverFormatoGalenhos 'Format(Now, sighEntidades.DevuelveHoraSoloFormato_HM)
        
End Sub
Sub CargarDatosParaEnviarAHospitalizacion()
        
        CargarDatosDeLasAtencionesPadres
        
        '1ro:   CARGAR DATOS DEL PACIENTE

        
        
        
        ml_ldFechaIngreso = lcBuscaParametro.RetornaFechaServidorSQL 'Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY)
        ml_lcHoraIngreso = lcBuscaParametro.RetornaHoraServidorSQLserverFormatoGalenhos 'Format(Now, sighEntidades.DevuelveHoraSoloFormato_HM)

End Sub

Sub DeshabilitarControlesParaEdicion()
    
    fraDatosReferenciaDestino.Enabled = False
    

End Sub

Sub ValoresPorDefecto()

    ml_ldFechaIngreso = lcBuscaParametro.RetornaFechaServidorSQL
    ml_lcHoraIngreso = lcBuscaParametro.RetornaHoraServidorSQLserverFormatoGalenhos
    Me.ucNacimientoDetalle1.FechaIngreso = CDate(Format(ml_ldFechaIngreso, sighEntidades.DevuelveFechaSoloFormato_DMY) & " " & ml_lcHoraIngreso)
    
End Sub


Sub Form_Load()
    If mo_lbCargaTablasUnaVez = True Then
        lbTieneLicenciaParaMensajeAcelulares = mo_sighProxies.VerificaLicenciaMensajeTexto
        lbCargaTablasUnaVez = False
        
        InicilizarParametros
        
        Me.ucDiagnosticosIngreso.Inicializar
        Me.ucDiagnosticoComplicaciones.Inicializar
        Me.ucDiagnosticoNacimiento.Inicializar
        Me.ucDiagnosticosEgreso.Inicializar
        Me.ucDiagnosticosMortalidad.Inicializar
        Me.ucNacimientoDetalle1.Inicializar
        Me.ucTransferenciasDetalle1.Inicializar
        CargarComboBoxes
        '
        lbBuscaDNIenReniec = IIf(wxParametro296 = "S", True, False)
        If lbBuscaDNIenReniec = True Then
           mo_Reniec.SeAccesaAlaWebDesdeGalenhos = True
           mo_Reniec.Inicializar
        End If
        '

    End If
    '
    SiempreCargaPorMovimiento
End Sub

Sub SiempreCargaPorMovimiento()
    If mo_lbNuevoMovimiento = True Then
        If Val(wxParametro208) = 1910 Then   'sullana
           lblGravedad.Caption = "Prioridad"
        End If
       
       lnIdTipoGravedadEgreso = 0
       If ml_TipoServicio = sghEmergenciaConsultorios Then
          mo_cmbServicioDestino.BoundColumn = "IdServicio"
          mo_cmbServicioDestino.ListField = "DservicioHosp"
          Set mo_cmbServicioDestino.RowSource = mo_AdminAdmision.DevuelveServiciosDelHospital("(3)", "", sghFiltraSoloActivos, sghPorDescServicio)
       Else
          cmbServicioDestino.Visible = False
       End If
        mo_lbNuevoMovimiento = False
        lbPacienteNN = False
        lcCaptionTab2 = ""
        lnFocusCuandoCargeFrm = 0
        lnIdNacimientoSeleccionado = 0
        '
        Select Case ml_TipoServicio
        Case sghHospitalizacion
        Case sghEmergenciaConsultorios
             cmbServicioDestino.Visible = True
             lblServicioDestino.Visible = True
        Case sghEmergenciaObservacion
        End Select
        '
        '
        btnAceptar.Enabled = True: btnAceptar.Visible = True
        '
        LimpiaTodosControles
        ConfiguraTABSsegunPermisosDelUsuario
        ConfigurarControles
        CargarDatosAlFormulario
        mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
        '
        
    End If
End Sub

Sub LimpiaTodosControles()
    If mi_Opcion = sghAgregar Then
            '
            '
            cmbIdDestinoAtencion.Text = ""
            txtIdServicioEgreso.Text = ""
            txtIdServicioEgreso.Tag = ""
            lblNombreServicioEgreso.Text = ""
            lblNombreMedicoEgreso.Text = ""
            txtIdMedicoEgreso.Text = ""
            cmbTipoAlta.Text = ""
            cmbCondicionAlta.Text = ""
            txtNroCamaEgreso.Text = ""
            txtFechaEgreso.Text = sighEntidades.FECHA_VACIA_DMY
            txtHoraEgreso.Text = sighEntidades.HORA_VACIA_HM
            txtFechaEgresoAdm.Text = sighEntidades.FECHA_VACIA_DMY
            txtHoraEgresoAdm.Text = sighEntidades.HORA_VACIA_HM
            cmbIdTipoReferenciaDestino.Text = ""
            txtIdEstablecimientoDestino.Text = ""
            lblNombreDestinoReferencia.Text = ""
            txtReferenciaD.Text = ""
            cmbServicioReferenciaD.Text = ""            'debb-21/06/2016
            txtFextension.Text = sighEntidades.FECHA_VACIA_DMY  'debb-21/06/2016
            txtFtramite.Text = sighEntidades.FECHA_VACIA_DMY    'debb-21/06/2016
            txtIdMedicoNacimiento.Text = ""
            lblNombreMedicoNacimiento.Text = ""
            chkSeRealizoNecropsia.Value = 0
            Me.txtIdMedicoEgreso.Tag = ""
            Me.txtIdMedicoNacimiento.Tag = ""
            Me.txtIdEstablecimientoDestino.Tag = ""
            Me.txtNroCamaEgreso.Tag = ""
            Me.txtIdServicioEgreso.Tag = ""
            '
            mo_cmbIdEspecialidadMedico.BoundText = ""
            mo_cmbIdServicio.BoundText = ""
            mo_cmbIdDestinoAtencion.BoundText = ""
            mo_cmbIdTipoReferenciaDestino.BoundText = ""
            mo_cmbCondicionAlta.BoundText = ""
            mo_cmbTipoAlta.BoundText = ""
            '
            lnIdDistritoSIS = 0: lnIdSexoSIS = 0: ldFechaNacimientoSIS = 0: lcSnombreSIS = "": lnIdPlanSIS = 0
            '
            ml_EstadoCuenta = 0
            ml_idCuentaAtencion = 0
            ml_idAtencion = 0
            ml_IdAtencionPadre = 0
            mo_NroServiciosQuePasoElPaciente = 0
            ml_IdPaciente = 0
            ms_Autogenerado = ""
            ml_IdMedico = 0
            ms_NombreMedico = ""
            ml_IdPrestamo = 0
            lcApP = ""
            lcApM = ""
            lcPnom = ""
            lcSnombreReniec = "": ldFnacimientoReniec = 0: lnIdSexoReniec = 0: lcDireccionReniec = "": mb_UsoWebReniec = False
            ml_IdServicioConCamaDisponible = 0
            lbPacienteNN = False
            lnFocusCuandoCargeFrm = 0
            lnIdPlanAnterior = 0
            lnIdTipoFinanciamientoAnterior = 0
    End If
    lcUltimoCodigoDeServicioTransferido = ""
    txtNroCamaEgreso.Visible = True: lblNroCamaEgreso.Visible = True
    lblFechaEgresoAdm.Visible = True: txtFechaEgresoAdm.Visible = True: txtHoraEgresoAdm.Visible = True
    
    '
    Me.ucDiagnosticoComplicaciones.LimpiarDatos
    Me.ucDiagnosticoNacimiento.LimpiarDatos
    Me.ucDiagnosticosEgreso.LimpiarDatos
    Me.ucDiagnosticosMortalidad.LimpiarDatos
    Me.ucNacimientoDetalle1.LimpiarDatos
    '
    
    
    
    
    
    '
    Set mo_Diagnosticos = Nothing
    
End Sub


Sub ConfiguraTABSsegunPermisosDelUsuario()
    
    Dim oRsPermisosTabs As New Recordset
    Me.TabEgresos.TabsPerRow = 4
    Me.TabEgresos.TabVisible(0) = True
    Me.TabEgresos.TabVisible(1) = False
    Me.TabEgresos.TabVisible(2) = False
    Me.TabEgresos.TabVisible(3) = False
    Set oRsPermisosTabs = ms_ReglasSeguridad.UsuariosRolesSeleccionarPermisosTodos(ml_idUsuario)
    If oRsPermisosTabs.RecordCount > 0 Then
       Do While Not oRsPermisosTabs.EOF
          Select Case oRsPermisosTabs.Fields!IdPermiso
          Case 350    'Admision Hosp/Emerg - Ver TAB 1.   Datos del Paciente
          Case 351    'Admisión Hosp/Emerg - Ver TAB 2.1 Ingreso
          Case 352    'Admisión Hosp/Emerg - Ver TAB 2.2 Transferencia
          Case 353    'Admisión Hosp/Emerg - Ver TAB 2.3 Causas Externas morbilidad Ing.
          Case 354    'Admisión Hosp/Emerg - Ver TAB 2.4 Diagnósticos Ing.
          Case 355    'Admisión Hosp/Emerg - Ver TAB 3.1 Egreso
          Case 356    'Admisión Hosp/Emerg - Ver TAB 3.2 Dx y Complicaciones Egr
               Me.TabEgresos.TabVisible(1) = True
          Case 357    'Admisión Hosp/Emerg - Ver TAB 3.3 Nacimientos Egr
               Me.TabEgresos.TabVisible(2) = True
          Case 358    'Admisión Hosp/Emerg - Ver TAB 3.4 Mortalidad Egr
               Me.TabEgresos.TabVisible(3) = True
          Case 362    'Admisión Hosp - Confirmar llegada de Paciente desde Adm.Emerg
               lbUsuarioConfirmaLlegada = True
          Case 363    'Admisión Hosp - Confirmar llegada de Paciente Transferido
               lbUsuarioConfirmaTransferencia = True
          End Select
          oRsPermisosTabs.MoveNext
       Loop
    End If
    Set oRsPermisosTabs = Nothing
   ' If ml_TipoServicio <> sghHospitalizacion Then
   '    Me.TabEgresos.TabVisible(2) = False
   ' End If
End Sub
Sub ConfigurarControles()
Dim oDOCuentaAtencionPadre As New DOCuentaAtencion
Dim oConexion As New Connection
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
     
    Me.ucDiagnosticoComplicaciones.TituloFrame = "Complicaciones     (F1=Todos Dx)"
    Me.ucDiagnosticosEgreso.TituloFrame = "Diagnósticos egreso     (F1=Todos Dx)"
    HabilitarFrameDestino False

    Select Case ml_TipoServicio
    Case sghHospitalizacion
        
        TituloDeForm "hospitalización"
    
        
                
                
        
                
                
                
    Case sghEmergenciaConsultorios
        
        TituloDeForm "emergencia"
        
        txtNroCamaEgreso.Visible = False: lblNroCamaEgreso.Visible = False
        lblFechaEgresoAdm.Visible = False: txtFechaEgresoAdm.Visible = False: txtHoraEgresoAdm.Visible = False
        
    
    
    
    Case sghEmergenciaObservacion
    
        'No se debe ver nacimientos
        TabEgresos.TabVisible(2) = False    'TabEgresos.TabVisible(1) = False
        TabEgresos.TabsPerRow = 3
        
        TituloDeForm "emergencia"
        
        
        
        
        
    End Select

    oConexion.Close
    Set oConexion = Nothing
End Sub
Sub TituloDeForm(sTitulo As String)
        
        Select Case mi_Opcion
        Case sghAgregar
            Me.Caption = "Agrega admisión " & sTitulo
        Case sghModificar
            Me.Caption = "Alta Médica " & sTitulo
        Case sghConsultar
            Me.Caption = "Consulta admisión " & sTitulo
        Case sghEliminar
            Me.Caption = "Elimina admisión " & sTitulo
        End Select

End Sub
'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla CuentasAtencion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Activate()
   If lcBuscaParametro.RetornaHoraServidorSQL < wxHoraMadrugada Then
       mo_lbCargaTablasUnaVez = True
   End If
   SiempreCargaPorMovimiento
   If mi_Opcion <> sghAgregar Then
        If Not mb_ExistenDatos Then
           Me.Visible = False
           LimpiarVariablesDeMemoria
        End If
        On Error Resume Next
        Select Case lnFocusCuandoCargeFrm
        Case 0  'No se ha realizado ninguna transferencia
        Case 1   'Confirmacion de Transferencias
        Case 2   'confirmacion de llegada al Servicio
        Case 3   'al  menos se realizó una transferencia
             lcCaptionTab2 = TabEgresos.Caption
             If TabEgresos.Tab = 0 Then
                cmbIdDestinoAtencion.SetFocus
             End If
        End Select
        lnFocusCuandoCargeFrm = 100
   Else
        If ml_TipoServicio = sghEmergenciaConsultorios Then
        End If
   End If
End Sub



Sub AdministrarKeyPreview(KeyCode As Integer)
    
    Select Case KeyCode
    'Case vbKeyEscape
    '    btnCancelar_Click
    Case vbKeyF2
        btnAceptar_Click
    End Select
       
End Sub

Sub ImprimeFormularioEnPDF()
       On Error GoTo errImpfrom
       Dim oReglasCaja As New SIGHNegocios.ReglasCaja
       Dim oConexion As New Connection
       Dim lbSePuedeImprimirPDF As Boolean, lcArchivoPDF As String
       Dim lcArchivoPDF0 As String
       sighEntidades.AbreConexionSIGH oConexion
       
       lcArchivoPDF0 = sighEntidades.DevuelveRutaConSlashInvertida(lcBuscaParametro.SeleccionaFilaParametro(237)) & _
       Trim(Str(mo_Atenciones.idPaciente)) & "-CTA" & Trim(Str(mo_Atenciones.idCuentaAtencion)) & "-GRABA-" & _
       Format(Date, "DDMMYYYY") & "-ALTA-" & _
       UCase(oReglasCaja.SeleccionaDatosCajeroConexion(sighEntidades.Usuario, sghUsuario, oConexion)) & "-Hr" & Format(Now, "hhmmss")
       oConexion.Close
       
       Dim lnFor As Integer
       For lnFor = 0 To 3
            TabEgresos.Tab = lnFor
            lcArchivoPDF = lcArchivoPDF0 & "-" & Mid(TabEgresos.Caption, 5, 4) & ".PDF"
            If SePuedeImprimirPDF(lcArchivoPDF, False) = True Then
                 lbSePuedeImprimirPDF = True
                 Me.PrintForm
            End If
       Next
       
       
'       lcArchivoPDF = lcArchivoPDF0 & "-Egreso.PDF"
'       If SePuedeImprimirPDF(lcImpresoraDefaultActual, lcArchivoPDF, False) = True Then
'            lbSePuedeImprimirPDF = True
'            TabEgresos.Tab = 0
'            Me.PrintForm
'       End If
'
'        lcArchivoPDF = lcArchivoPDF0 & "-Complicaciones.PDF"
'        If SePuedeImprimirPDF(lcImpresoraDefaultActual, lcArchivoPDF, False) = True Then
'            TabEgresos.Tab = 1
'            Me.PrintForm
'        End If
'
'        lcArchivoPDF = lcArchivoPDF0 & "-Nacimientos.PDF"
'        If SePuedeImprimirPDF(lcImpresoraDefaultActual, lcArchivoPDF, False) = True Then
'            TabEgresos.Tab = 2
'            Me.PrintForm
'        End If
'
'        lcArchivoPDF = lcArchivoPDF0 & "-Mortalidad.PDF"
'        If SePuedeImprimirPDF(lcImpresoraDefaultActual, lcArchivoPDF, False) = True Then
'            TabEgresos.Tab = 3
'            Me.PrintForm
'        End If
        
        
        SeteaOtraImpresoraDefault sighEntidades.ImpresoraDefaultDeEstaPC
errImpfrom:
       
       Set oReglasCaja = Nothing
       Set oConexion = Nothing
End Sub



Private Sub btnAceptar_Click()
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   Dim oConexion As New Connection
   oConexion.Open sighEntidades.CadenaConexion
   oConexion.CursorLocation = adUseClient
   Select Case mi_Opcion
   Case sghModificar
       If ValidarDatosObligatorios() Then
            CargaDatosAlObjetosDeDatos False
           If ValidarReglas() Then
               If Not ValidarDiasInternamiento() Then Exit Sub
               If ModificarDatos() Then
                   
                   'ImprimeFormularioEnPDF
                   
                   If lbTieneLicenciaParaMensajeAcelulares = True Then
                        Dim oMensajeCelular As New SIGHProxies.Procesos
                        oMensajeCelular.MensajeCelularEnviarSegunDxGRAVE mo_Pacientes, Me.ucDiagnosticosEgreso.DevuelveDx, _
                                                                         IIf(ml_TipoServicio <> sghHospitalizacion, "ALTA EMERGENCIA", "ALTA HOSPITALIZACION"), _
                                                                         mo_Atenciones.idCuentaAtencion, oConexion
                        Set oMensajeCelular = Nothing
                   End If
                   If lbSeEnviaMensajeCelularAreaConv = True Then
                      
                        Dim oMensajeCelular1 As New SIGHProxies.Procesos
                        oMensajeCelular1.MensajeCelularEnviarAconvenios mo_Pacientes, "CONV", _
                                        mo_Atenciones.idCuentaAtencion, oConexion, _
                                        lblNombreServicioEgreso.Text
                        Set oMensajeCelular1 = Nothing
                   End If
                   MsgBox " Los datos se modificaron correctamente, para la Cuenta N° " & Trim(Str(ml_idCuentaAtencion)), vbInformation, Me.Caption
                   If wxParametro302 = "S" And mo_Atenciones.IdFormaPago = 2 And Me.txtFechaEgreso.Text <> sighEntidades.FECHA_VACIA_DMY Then
                        'mgaray201410e
                        'RZC 28/12/2020 Cambio por error de impresión de FUA en hospitalizacion INICIO
                        'Se comento:
                        'If ServicioImprimeFUAAdmision() = False Then
                        '    btnImprimeFichaSIS_Click
                        'End If
                        'RZC 28/12/2020 Cambio por error de impresión de FUA en hospitalizacion FIN
                   End If
                   If mo_Atenciones.IdDestinoAtencion = 21 And wxParametro552 = "S" Then
                      'es una ALTA DE EMERGENCIA CON DESTINO=HOSPITALIZACION
                      'Se creará una CUENTA DE HOSPITALIZACION en forma automática
                      CreaCuentaAutomaticaEnHospitalizacion
                   End If
                   Me.Visible = False
                   LimpiarVariablesDeMemoria
               Else
                   MsgBox "No se pudo modificar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   End Select
   oConexion.Close
   Set oConexion = Nothing
End Sub


Function ValidarDiasInternamiento() As Boolean
    ValidarDiasInternamiento = False

    ValidarDiasInternamiento = True
End Function

Private Sub btnCancelar_Click()
   Me.Visible = False
   LimpiarVariablesDeMemoria
End Sub

Function ValidarDatosObligatorios() As Boolean
   Dim sMensaje As String
   sMensaje = ""
   ValidarDatosObligatorios = False
   
   '-------------------------------------------------------------------------
   '                VALIDA DATOS DE LA CUENTA DE ATENCION
   '-------------------------------------------------------------------------
   
    '---------------------------------------------------------------------------------
    '           VALIDA DATOS DE LA ATENCION
    '---------------------------------------------------------------------------------
   '
   CambioFechaNacimiento Format(mo_Pacientes.FechaNacimiento, sighEntidades.DevuelveFechaSoloFormato_DMY), _
                         Format(mo_Pacientes.FechaNacimiento, sighEntidades.DevuelveHoraSoloFormato_HM)
   
   
   '
   '---------------------------------------------------------------------------------
   '           VALIDA DATOS DE PACIENTES
   '---------------------------------------------------------------------------------
   '
   If wxParametro530 = "S" And mo_cmbIdDestinoAtencion.BoundText = "21" And cmbServicioDestino.Text = "" Then
      sMensaje = sMensaje & "Tiene que elegir el SERVICIO DEST" & Chr(13)
   End If
   
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
    Dim lcMensaje As String
    Dim rsCitas As New Recordset
    ValidarReglas = False
    
    If wxParametro552 = "S" Then
        If Not mo_AdminAdmision.ValidaEdadMaximaYSexoSegunServicioHosp(mo_Atenciones.Edad, mo_Atenciones.idTipoEdad, _
                                    mo_Pacientes.idTipoSexo, Val(mo_cmbServicioDestino.BoundText), True) Then
            Exit Function
        End If
    End If
   
    Dim lIdCausaBasica As Long
    Dim lIdCausaIntermedia As Long
    Dim lIdCausaFinal As Long
    Dim lIdDxPrincipal As Long
    Dim lIdDxIngreso As Long
    
    ObtenerDiagnosticos lIdDxPrincipal, lIdCausaBasica, lIdCausaIntermedia, lIdCausaFinal, lIdDxIngreso
      
    If Me.txtFechaEgreso <> sighEntidades.FECHA_VACIA_DMY Then
    
    
         If CDate(Me.txtFechaEgreso) > lcBuscaParametro.RetornaFechaServidorSQL Then
            MsgBox "No puede dar alta con FECHA mayor a HOY", vbInformation, Me.Caption
            Exit Function
         End If
    
         If txtIdMedicoEgreso.Text = "" Then
            MsgBox "Por favor debe registrar el Médico del Alta", vbExclamation, Me.Caption
            Exit Function
         End If
         If ml_TipoServicio <> sghHospitalizacion Then
            '****emergencia *****
            If wxParametro552 = "S" And mo_cmbIdDestinoAtencion.BoundText = "21" Then
               If cmbServicioDestino.Text = "" Then
                    MsgBox "Por favor debe elegir SERVICIO DESTINO", vbInformation, ""
                    Exit Function
               End If
            End If
            
            If mo_cmbIdDestinoAtencion.BoundText = "25" Then
               mo_cmbTipoAlta.BoundText = "7"
               mo_cmbCondicionAlta.BoundText = "4"
            ElseIf mo_cmbTipoAlta.BoundText = "7" Then
               MsgBox "El Paciente no falleció, eligió mal el TIPO DE ALTA", vbInformation, ""
               Exit Function
            ElseIf mo_cmbCondicionAlta.BoundText = "4" Then
               MsgBox "El Paciente no falleció, eligió mal el CONDICION DE ALTA", vbInformation, ""
               Exit Function
            End If
            '
            If lIdDxIngreso = 0 Then   'debb-05/04/2011
               MsgBox "Por favor asigne el Diagnóstico de INGRESO", vbInformation, Me.Caption
               Exit Function
            End If
            If DateDiff("h", Format(ml_ldFechaIngreso, sighEntidades.DevuelveFechaSoloFormato_DMY) & " " & ml_lcHoraIngreso, Me.txtFechaEgreso & " " & Me.txtHoraEgreso.Text) > Val(wxParametro324) Then
                MsgBox "En Emergencia, la FECHA DE EGRESO no  puede pasar de " & wxParametro324 & " horas", vbExclamation, Me.Caption
                Exit Function
            End If
            '
            Dim mo_ReglasServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
            Dim oDoServicio As New doServicio
            Dim oConexion As New Connection
            oConexion.CommandTimeout = 900
            oConexion.CursorLocation = adUseClient
            oConexion.Open sighEntidades.CadenaConexion
            Set oDoServicio = mo_ReglasServiciosHosp.ServiciosSeleccionarPorId(Me.txtIdServicioEgreso.Tag, oConexion)
            If oDoServicio.EsObservacionEmergencia = True And txtNroCamaEgreso.Text = "" Then
               oConexion.Close
               Set oDoServicio = Nothing
               Set mo_ReglasServiciosHosp = Nothing
               MsgBox "El servicio es OBSERVACION DE EMERGENCIA, falta asignar CAMA", vbInformation, Me.Caption
               Exit Function
            End If
            oConexion.Close
            Set oDoServicio = Nothing
            Set mo_ReglasServiciosHosp = Nothing
            
         Else
            If mo_cmbIdDestinoAtencion.BoundText = "35" Or mo_cmbIdDestinoAtencion.BoundText = "33" Then
               mo_cmbTipoAlta.BoundText = "7"
               mo_cmbCondicionAlta.BoundText = "4"
            ElseIf mo_cmbTipoAlta.BoundText = "7" Then
               MsgBox "El Paciente no falleció, eligió mal el TIPO DE ALTA", vbInformation, ""
               Exit Function
            ElseIf mo_cmbCondicionAlta.BoundText = "4" Then
               MsgBox "El Paciente no falleció, eligió mal el CONDICION DE ALTA", vbInformation, ""
               Exit Function
            End If
            '
            'nacimientos en Hospitalizacion
            Dim oDOAtencionNacimiento As New DOAtencionNacimiento
            Dim oDOAtencionDiagnostico As New DOAtencionDiagnostico
            Dim lbContinuar As Boolean
            If mo_Nacimientos.Count > 0 Then
                lbContinuar = True
                '
'                If mo_Pacientes.IdDocIdentidad <> 1 Then
'                   lbContinuar = False
'                ElseIf Len(mo_Pacientes.NroDocumento) <> 8 Then
'                   lbContinuar = False
'                End If
'                If lbContinuar = False Then
'                    MsgBox "(Ficha 1) debe ingresar el DNI de la madre (8 digitos) ", vbExclamation, Me.Caption
'                    Exit Function
'                End If
                '
                If lbContinuar = True Then
                    For Each oDOAtencionNacimiento In mo_Nacimientos
                        If oDOAtencionNacimiento.idCondicionRN = 2 Then
                            lbContinuar = False
                            If Not mo_Diagnosticos Is Nothing Then
                                For Each oDOAtencionDiagnostico In mo_Diagnosticos
                                    If oDOAtencionDiagnostico.IdClasificacionDx = sghHospitalizacionNacimiento Then
                                       lbContinuar = True
                                       Exit For
                                    End If
                                 Next
                            End If
                            If lbContinuar = False Then
                                MsgBox "Ficha (3.3) Nacimientos -> Hubo un nacido MUERTO, deberá ingresar un 'Dx de Muerte Fetal'", vbExclamation, Me.Caption
                                Exit Function
                            Else
                                Exit For
                            End If
                        End If
                     Next
                 End If
            End If
            Set oDOAtencionNacimiento = Nothing
            Set oDOAtencionDiagnostico = Nothing
         End If
    
         If CDate(Me.txtFechaEgreso.Text & " " & Me.txtHoraEgreso.Text) < CDate(Format(ml_ldFechaIngreso, sighEntidades.DevuelveFechaSoloFormato_DMY) & " " & ml_lcHoraIngreso) Then
             MsgBox "La fecha de EGRESO MEDICO no puede ser menor que la fecha de INGRESO", vbExclamation, Me.Caption
             Exit Function
         End If
        
         If DateDiff("d", Format(ml_ldFechaIngreso, sighEntidades.DevuelveFechaSoloFormato_DMY), Me.txtFechaEgreso) > Val(sighEntidades.EstanciaMaxHospitalizacion) Then
             If MsgBox("¡El intervalo entre la fecha de ingreso y egreso es de mayor que la estancia máxima (" & _
                         sighEntidades.EstanciaMaxHospitalizacion + " días) ¡" & Chr(13) & "Fecha de Ingreso: " & _
                         ml_ldFechaIngreso & " -  Fecha de Egreso: " & Me.txtFechaEgreso & Chr(13) & _
                         "¿Es correcto?", vbYesNo + vbQuestion, Me.Caption) = vbNo Then
                Exit Function
            End If
         End If
    
         If Me.txtFechaEgresoAdm <> sighEntidades.FECHA_VACIA_DMY Then
            If CDate(Me.txtFechaEgresoAdm) < CDate(Me.txtFechaEgresoAdm) Then
                MsgBox "La fecha de egreso administrativo no puede ser menor que la fecha de egreso médico", vbExclamation, Me.Caption
                Exit Function
            End If
         End If
        
         If Val(mo_cmbCondicionAlta.BoundText) = 4 Then   'Paciente Fallecido
            If lIdCausaFinal = 0 Then
                MsgBox "La condición del paciente indica PACIENTE FALLECIDO, Por favor ingreso de DIAGNOSTICO -> CAUSA FINAL (ficha 3.4)", vbInformation, Me.Caption
                Me.TabEgresos.Tab = 3
                ucDiagnosticosEgreso.LimpiarDatos
                Exit Function
            End If
            If lIdCausaBasica = 0 Then
                If lIdCausaIntermedia <> 0 Then
                    MsgBox "Por favor antes de llenar la causa intermedia debe llenar la causa básica", vbInformation, Me.Caption
                    Exit Function
                End If
            End If
         Else
           If lblNombreMedicoEgreso.Text <> "" Then
              If lIdDxPrincipal = 0 Then
                  MsgBox "Por favor debe llenar DIAGNOSTICO -> PRINCIPAL", vbInformation, Me.Caption
                  Exit Function
              End If
              If cmbCondicionAlta.Text = "" Then
                 MsgBox "Por favor debe elegir la CONDICION DE ALTA", vbInformation, Me.Caption
                 Exit Function
              End If
           End If
         End If
         '
         If mo_ReglasSISgalenhos.SisFUAyaFueEnviadoAlSisLIMA(ml_idCuentaAtencion, ml_IdFuenteFinanciamiento, wxParametro302) = True Then
            Exit Function
         End If
         If mo_Pacientes.idTipoSexo = 2 And DestinoAntecionChequeaEnPartos = True Then
            If mo_AdminAdmision.ChequeaSiTuboDxParto(mo_Diagnosticos) = True Then
               If mo_Nacimientos.Count = 0 Then
                    MsgBox "La paciente tubo Dx de PARTO, por favor debe registrar cada Nacimiento (Ficha 3.3)" & _
                            Chr(13) & Chr(13) & "(según DESTINO ATENCION)", vbInformation, Me.Caption
                    On Error Resume Next
                    TabEgresos.Tab = 2
                    Exit Function
               End If
            End If
         End If
    End If
    '
    
    'Verifica si algunos de los servicios es de cirugia, ginecologia u obstetricia
    Dim bServPerteneceACirugiaOGinecologia As Boolean
    Dim bServicioPerteneceAPediatria As Boolean
    bServPerteneceACirugiaOGinecologia = False
    bServicioPerteneceAPediatria = False
    Dim oDOOcupacion As New DOEstanciaHospitalaria
    Dim lIdDepartamento As Long
    
    
    For Each oDOOcupacion In mo_OcupacionCamas
        lIdDepartamento = mo_AdminServiciosHosp.ServiciosSeleccionarIdDepartamento(oDOOcupacion.IdServicio)
        If lIdDepartamento = 3 Or lIdDepartamento = 4 Then  'Cirugia y (Ginecologia y Obstetricia)
            bServPerteneceACirugiaOGinecologia = True
        End If
        If lIdDepartamento = 2 Then  'Pediatria
            bServicioPerteneceAPediatria = True
        End If
        
    Next
    If mi_Opcion = sghModificar Then
        If mo_NroServiciosQuePasoElPaciente < mo_OcupacionCamas.Count And lcUltimoCodigoDeServicioTransferido = Me.txtIdServicioEgreso.Text Then
           MsgBox "No puede transferir al mismo SERVICIO", vbInformation, Me.Caption
           Exit Function
        End If
    End If
    
    'Si tiene algun servicio del dpto de cirugia , ginecologia y obstetricia
    If bServPerteneceACirugiaOGinecologia Then
        If mo_Procedimientos.Count = 0 Then
            

        End If
    End If
    
    'Valida la infeccion intrahospitalaria
    If Not HuboDiagnosticoInfeccionAlIngreso() Then
        If HayDiagnosticosInfeccionAlEgreso() Then
            If MsgBox("Los diagnosticos muestran la existencia de infección intrahospitalaria, ¿Es correcto?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                mo_DoAtencionDatosAdicionales.HuboInfeccionIntraHospitalaria = False
            Else
                mo_DoAtencionDatosAdicionales.HuboInfeccionIntraHospitalaria = True
            End If
        End If
    End If
    
    'referencia EGRESO
    If cmbIdDestinoAtencion.Text <> "" Then
         Dim sCodigoDestino As String
        sCodigoDestino = Trim(Split(cmbIdDestinoAtencion.Text, " = ")(0))
        If sCodigoDestino = "R" Or sCodigoDestino = "C" Then
           If lblNombreDestinoReferencia.Text = "" Then
              MsgBox "Debe elejir el ESTABLEC.REFERIDO (Destino)(ficha 3.1)", vbInformation, Me.Caption
              Me.TabEgresos.Tab = 0
              Exit Function
           End If
           If txtReferenciaD.Text = "" Then
              MsgBox "Debe registrar el N°REFERENCIA (Destino)(ficha 3.1)", vbInformation, Me.Caption
              Me.TabEgresos.Tab = 0
              Exit Function
           End If
           If mo_AdminServiciosComunes.BuscarSiExisteNUMEROREFERENCIA(txtReferenciaD.Text, ml_idAtencion, IIf(sCodigoDestino = "C", True, False)) = True Then
              Me.TabEgresos.Tab = 0
              txtReferenciaD.Text = mo_AdminServiciosComunes.CalculaNUMEROREFERENCIA(IIf(sCodigoDestino = "C", True, False))
              Exit Function
           End If
           If cmbServicioReferenciaD.Text = "" Then
              MsgBox "Debe elegir el SERVICIO DE LA REFERENCIA (ficha 2.1)", vbInformation, Me.Caption
              Me.TabEgresos.Tab = 0
              Exit Function
           End If
           '
           If Me.txtFextension.Text = sighEntidades.FECHA_VACIA_DMY Then
                    MsgBox "Debe registrar la fecha de EXTENSION DE LA REFERENCIA  (ficha 2.1)", vbInformation, Me.Caption
                    Me.TabEgresos.Tab = 0
                    Exit Function
           Else
              If IsDate(Me.txtFextension.Text) Then
                 If Me.txtFextension.Text <> sighEntidades.FECHA_VACIA_DMY And Me.txtFtramite.Text <> sighEntidades.FECHA_VACIA_DMY Then
                    If CDate(Me.txtFextension.Text) > CDate(Me.txtFtramite.Text) Then
                          MsgBox "La fecha de EXTENSION DE LA REFERENCIA no puede ser mayor a la  FECHA DE TRAMITE (ficha 2.1)", vbInformation, Me.Caption
                          Me.TabEgresos.Tab = 0
                          Exit Function
                    End If
                 End If
                 If CDate(Me.txtFextension.Text) < ml_ldFechaIngreso Then
                      MsgBox "La fecha de EXTENSION DE LA REFERENCIA no puede ser menor a la FECHA DE INGRESO (ficha 2.1)", vbInformation, Me.Caption
                      Me.TabEgresos.Tab = 0
                      Exit Function
                 End If
              
              Else
                      MsgBox "La fecha de EXTENSION DE LA REFERENCIA no es VALIDA (ficha 2.1)", vbInformation, Me.Caption
                      Me.TabEgresos.Tab = 0
                      Exit Function
              End If
           End If
           '
           If Me.txtFtramite.Text = sighEntidades.FECHA_VACIA_DMY Then
                    MsgBox "Debe registrar la fecha de TRAMITE DE LA REFERENCIA  (ficha 2.1)", vbInformation, Me.Caption
                    Me.TabEgresos.Tab = 0
                    Exit Function
           Else
              If IsDate(Me.txtFtramite.Text) Then
                 If CDate(Me.txtFtramite.Text) < ml_ldFechaIngreso Then
                      MsgBox "La fecha de TRAMITE DE LA REFERENCIA no puede ser menor a la FECHA DE INGRESO (ficha 2.1)", vbInformation, Me.Caption
                      Me.TabEgresos.Tab = 0
                      Exit Function
                 End If
              Else
                      MsgBox "La fecha de TRAMITE DE LA REFERENCIA no es VALIDA (ficha 2.1)", vbInformation, Me.Caption
                      Me.TabEgresos.Tab = 0
                      Exit Function
              End If
           End If
           '
        End If
    End If
    
    
    If mi_Opcion = sghAgregar And wxParametro302 = "S" And ml_IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
       lcMensaje = mo_ReglasSISgalenhos.ChequeaCodigoEstablecimientoAdscripcion(lcCodigoEstablecimientoAdscripcionSIS, _
                                           ml_TipoServicio, _
                                           mo_AdminAdmision.TiposOrigenAtencionDevuelveIdSis(ml_IdViasAdmision), _
                                           "")
       If lcMensaje <> "" Then
             MsgBox lcMensaje, vbInformation, Me.Caption
             
             CargarAutomaticamenteEstablecimientoDestinoSIS ' Frank 2608
            
             Exit Function
       End If
    End If
    If mi_Opcion = sghModificar And Me.txtFechaEgreso.Text <> sighEntidades.FECHA_VACIA_DMY Then
       If Me.UcEpisodioClinico1.ValidaReglas(" (Ficha 3.1)") = False Then
          Me.TabEgresos.Tab = 0
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
    Set oDOOcupacion = Nothing
    Set rsCitas = Nothing
    ValidarReglas = True
End Function

Public Sub CargarAutomaticamenteEstablecimientoDestinoSIS() 'Frank 2808
                If lcBuscaParametro.SeleccionaFilaParametro(326) = "S" And lcCodigoEstablecimientoAdscripcionSIS <> "" Then
                   Dim lcCodigoSis As String
                   Dim lcEstablecimientoOrigen As String
                   Dim DOEstablecimiento As New DOEstablecimiento
                   Dim oRsEstabNoMINSA As Recordset
                   Dim lnIdOrigenDelPacienteDesdeFUA As Long
                   lnIdOrigenDelPacienteDesdeFUA = mo_AdminAdmision.TiposOrigenAtencionDevuelveIdSis(ml_IdViasAdmision)
                   
                   If Val(lcBuscaParametro.SeleccionaFilaParametro(280)) <> Val(lcCodigoEstablecimientoAdscripcionSIS) Then
                      If lcBuscaParametro.SeleccionaFilaParametro(282) <> "S" Then 'Hospital
                           If Not (lnIdOrigenDelPacienteDesdeFUA = "4" Or lnIdOrigenDelPacienteDesdeFUA = "6") Then 'Referido CE, ContraReferido
'                                ml_IdViasAdmision = 12
                                If mo_AdminServiciosComunes.EstablecimientosSeleccionarPorCodigo(Right(lcCodigoEstablecimientoAdscripcionSIS, 5), DOEstablecimiento) = True Then
                                    mo_cmbIdTipoReferenciaDestino.BoundText = 1 'MINSA
                                    txtIdEstablecimientoDestino.Text = DOEstablecimiento.Codigo
                                    txtIdEstablecimientoDestino.Tag = DOEstablecimiento.IdEstablecimiento
                                    lblNombreDestinoReferencia.Text = DOEstablecimiento.nombre
                                Else
                                    Set oRsEstabNoMINSA = mo_AdminServiciosComunes.EstablecimientosNoMinsaSeleccionarPorCodigo(Right(lcCodigoEstablecimientoAdscripcionSIS, 5))
                                    If oRsEstabNoMINSA.RecordCount > 0 Then
                                        oRsEstabNoMINSA.MoveFirst
                                        mo_cmbIdTipoReferenciaDestino.BoundText = 2 'NO MINSA
                                        txtIdEstablecimientoDestino.Text = oRsEstabNoMINSA.Fields!Codigo
                                        txtIdEstablecimientoDestino.Tag = oRsEstabNoMINSA.Fields!IdEstablecimientoNoMINSA
                                        lblNombreDestinoReferencia.Text = oRsEstabNoMINSA.Fields!nombre
                                    End If
                                    Set oRsEstabNoMINSA = Nothing
                                End If
                           End If
                      End If
                   End If
                   Set DOEstablecimiento = Nothing
                End If
End Sub

Function DestinoAntecionChequeaEnPartos() As Boolean
    DestinoAntecionChequeaEnPartos = False
    Select Case mo_cmbIdDestinoAtencion.BoundText
    Case "65", "20"                     'Emergencia: Citado/Domicilio/
        DestinoAntecionChequeaEnPartos = True
    Case "30", "52", "68"               'Hospitalizacion: Domicilio/SuCAsa/Citado
        DestinoAntecionChequeaEnPartos = True
    End Select
End Function


Function HayDiagnosticosInfeccionAlEgreso() As Boolean
Dim oDOAtencionDiagnostico As DOAtencionDiagnostico
Dim oDODiagnostico As DODiagnostico

    HayDiagnosticosInfeccionAlEgreso = False

    For Each oDOAtencionDiagnostico In mo_Diagnosticos
        If oDOAtencionDiagnostico.IdClasificacionDx = 3 Then    'Diagnostico de egreso
            Set oDODiagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorId(oDOAtencionDiagnostico.idDiagnostico)
            If oDODiagnostico.Intrahospitalario Then
                HayDiagnosticosInfeccionAlEgreso = True
            End If
        End If
    Next

End Function
Function HuboDiagnosticoInfeccionAlIngreso() As Boolean
Dim oDODiagnostico As New DODiagnostico
Dim oDOAtencionDiagnostico As DOAtencionDiagnostico

    HuboDiagnosticoInfeccionAlIngreso = False
    
    For Each oDOAtencionDiagnostico In mo_Diagnosticos
        If oDOAtencionDiagnostico.IdClasificacionDx = 2 Then    'Diagnostico de ingreso
            Set oDODiagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorId(oDOAtencionDiagnostico.idDiagnostico)
            If oDODiagnostico.Intrahospitalario Then
                HuboDiagnosticoInfeccionAlIngreso = True
            End If
        End If
    Next
    Set oDODiagnostico = Nothing

End Function
Function ObtenerDiagnosticos(lIdDxPrincipal As Long, lIdCausaBasica As Long, lIdCausaIntermedia As Long, lIdCausaFinal As Long, lIdDxIngreso)
    
    lIdDxPrincipal = 0
    lIdCausaBasica = 0
    lIdCausaIntermedia = 0
    lIdCausaFinal = 0
    lIdDxIngreso = 0
    
    Dim lnIdAtencion As Long, lnIdClasificacionDx As Long, lnIdSubClasificacionDx As Long, lnIdUsuario As Long
    Dim lcLabConfHIS As String, lnIdAtencionDiagnostico As Long
    
    Dim oDODiagnostico As DOAtencionDiagnostico
    For Each oDODiagnostico In mo_Diagnosticos
        If oDODiagnostico.IdSubclasificacionDx = 301 Then
           lnIdAtencionDiagnostico = oDODiagnostico.IdAtencionDiagnostico
           lnIdAtencion = oDODiagnostico.IdAtencionDiagnostico
           lIdDxPrincipal = oDODiagnostico.idDiagnostico
           lnIdClasificacionDx = oDODiagnostico.IdClasificacionDx
           lnIdSubClasificacionDx = oDODiagnostico.IdSubclasificacionDx
           lnIdUsuario = oDODiagnostico.IdUsuarioAuditoria
           lcLabConfHIS = oDODiagnostico.labConfHIS
        End If
        If oDODiagnostico.IdSubclasificacionDx = 303 Then lIdCausaFinal = oDODiagnostico.idDiagnostico
        If oDODiagnostico.IdSubclasificacionDx = 304 Then lIdCausaIntermedia = oDODiagnostico.idDiagnostico
        If oDODiagnostico.IdSubclasificacionDx = 305 Then lIdCausaBasica = oDODiagnostico.idDiagnostico
        If oDODiagnostico.IdSubclasificacionDx = 0 Then lIdDxIngreso = oDODiagnostico.idDiagnostico
    Next
    If lIdDxIngreso = 0 And (lIdDxPrincipal > 0 Or lIdCausaFinal > 0) And ml_TipoServicio = sghEmergenciaConsultorios Then  'debb-22/07/2016
        Set oDODiagnostico = New DOAtencionDiagnostico
        oDODiagnostico.IdAtencionDiagnostico = lnIdAtencionDiagnostico
        oDODiagnostico.idAtencion = lnIdAtencion
        oDODiagnostico.idDiagnostico = IIf(lIdCausaFinal > 0, lIdCausaFinal, lIdDxPrincipal)
        oDODiagnostico.IdClasificacionDx = 2
        oDODiagnostico.IdSubclasificacionDx = 0
        oDODiagnostico.IdUsuarioAuditoria = lnIdUsuario
        oDODiagnostico.labConfHIS = lcLabConfHIS
        mo_Diagnosticos.Add oDODiagnostico
        lIdDxIngreso = IIf(lIdCausaFinal > 0, lIdCausaFinal, lIdDxPrincipal)
    End If
End Function





'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla CuentasAtencion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Sub CargaDatosAlObjetosDeDatos(SeLlamaDesdeCreacionDeCuentaEnHospitalizacion As Boolean)
    'Limpia Dx
    Set mo_Diagnosticos = Nothing
    'Limpia Nacimientos
    Set mo_Nacimientos = Nothing
    'Limpia Transferencias
    Set mo_OcupacionCamas = Nothing
    '
    If txtHoraEgreso <> sighEntidades.HORA_VACIA_HM Then
       If mo_cmbCondicionAlta.BoundText = "4" Then
          'Si el paciente es "fallecido"  --> no debe tener DIAGNOSTICO EGRESO
          'solo Dx de MORTALIDAD
          ucDiagnosticosEgreso.LimpiarDatos
       Else
          'Si el paciente NO "fallece"  --> se debe tener DIAGNOSTICO EGRESO
          'eliminar los de MORTALIDAD
          ucDiagnosticosMortalidad.LimpiarDatos
       End If
    End If
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LA CUENTA ATENCION
        '---------------------------------------------------------------------------------
    Select Case ml_TipoAccionAdmision
    Case sghAdmisionNormal  'Si el una admisión normal de hospitalizacion o de emergencia
        Select Case mi_Opcion
        Case sghAgregar
            With mo_CuentasAtencion
                        .idCuentaAtencion = Me.idCuentaAtencion
                        .idPaciente = Me.idPaciente
                        .TotalAsegurado = 0
                        .TotalExonerado = 0
                        .TotalPagado = 0
                        .TotalPorPagar = 0
                        If ml_TipoServicio = sghHospitalizacion Then
                             If wxParametro232 = "S" Then
                                .idEstado = sghEstadoCuenta.sghNoLlegaAlServicioHospitalizado
                             Else
                                .idEstado = sghEstadoCuenta.sghAbierto
                             End If
                        Else
                             .idEstado = sghEstadoCuenta.sghAbierto   'Estado de la cuenta = ABIERTO
                        End If
                        'WCG 10/06
                        '.fechaApertura = IIf(Me.txtFechaIngreso.Text = sighentidades.HORA_VACIA_HM, "", Me.txtFechaIngreso.Text)
                        '.HoraApertura = IIf(Me.txtHoraIngreso.Text = sighentidades.HORA_VACIA_HM, "", Me.txtHoraIngreso.Text)
                        .fechaCierre = 0
                        .HoraCierre = ""
                        .IdUsuarioAuditoria = ml_idUsuario
            End With
        Case sghModificar
            '******Transferencias
            If lcUltimoCodigoDeServicioTransferido <> "" Then
               If ml_TipoServicio = sghEmergenciaConsultorios Then
                    Dim oDOServicioH1 As New doServicio
                    Set oDOServicioH1 = mo_AdminServiciosHosp.ServiciosSeleccionarPorCodigo(lcUltimoCodigoDeServicioTransferido)
                    If oDOServicioH1.EsObservacionEmergencia = True Then
                       mo_CuentasAtencion.idEstado = sghEstadoCuenta.sghNoLlegaAlServicioHospitalizado
                    Else
                       mo_CuentasAtencion.idEstado = 1
                    End If
                    Set oDOServicioH1 = Nothing
               Else
                    mo_CuentasAtencion.idEstado = sghEstadoCuenta.sghNoLlegaAlServicioHospitalizado
               End If
            End If
        End Select
    End Select
   
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LA ATENCION
    '---------------------------------------------------------------------------------
   With mo_Atenciones
           
           .IdDestinoAtencion = Val(mo_cmbIdDestinoAtencion.BoundText)
           
           .fechaEgreso = IIf(Me.txtFechaEgreso = sighEntidades.FECHA_VACIA_DMY, 0, Me.txtFechaEgreso)
           .HoraEgreso = IIf(Me.txtHoraEgreso = sighEntidades.HORA_VACIA_HM, "", Me.txtHoraEgreso)
           .IdUsuarioAuditoria = Me.idUsuario
               
   
            .FechaEgresoAdministrativo = IIf(Me.txtFechaEgresoAdm = sighEntidades.FECHA_VACIA_DMY, 0, Me.txtFechaEgresoAdm)
            .HoraEgresoAdministrativo = Me.txtHoraEgresoAdm
            .IdCamaEgreso = Val(Me.txtNroCamaEgreso.Tag)
            .IdCondicionAlta = Val(mo_cmbCondicionAlta.BoundText)
            If mi_Opcion = sghModificar Then
               If lcUltimoCodigoDeServicioTransferido <> "" Then
                    Dim oDOServicioH As New doServicio
                    Set oDOServicioH = mo_AdminServiciosHosp.ServiciosSeleccionarPorCodigo(lcUltimoCodigoDeServicioTransferido)
                    .IdServicioEgreso = oDOServicioH.IdServicio
                    .IdCamaEgreso = 0
                    Set oDOServicioH = Nothing
               Else
                   .IdServicioEgreso = Val(Me.txtIdServicioEgreso.Tag)
               End If
            Else
               .IdServicioEgreso = Val(Me.txtIdServicioEgreso.Tag)
            End If
               
            .IdTipoAlta = Val(mo_cmbTipoAlta.BoundText)
            .IdMedicoEgreso = Val(Me.txtIdMedicoEgreso.Tag)
            
'            If lnIdTipoGravedadEgreso > 0 Then
'               .IdTipoGravedad = lnIdTipoGravedadEgreso
'            End If
            .IdTipoGravedad = Val(mo_cmbIdTipoGravedad.BoundText)
   End With
   
   


    '---------------------------------------------------------------------------------
    '           CARGA DATOS DEL PACIENTE
    '---------------------------------------------------------------------------------

    '---------------------------------------------------------------------------------
    '           COMPLETA LOS DATOS DE LA ATENCION
    '---------------------------------------------------------------------------------
    With mo_DoAtencionDatosAdicionales
        .IdMedicoRespNacimiento = Val(Me.txtIdMedicoNacimiento.Tag)
        .IdTipoReferenciaDestino = Val(mo_cmbIdTipoReferenciaDestino.BoundText)
        If .IdTipoReferenciaDestino = 1 Then
             .idEstablecimientoDestino = Val(Me.txtIdEstablecimientoDestino.Tag)
             .IdEstablecimientoNoMinsaDestino = 0
        Else
             .idEstablecimientoDestino = 0
             .IdEstablecimientoNoMinsaDestino = Val(Me.txtIdEstablecimientoDestino.Tag)
        End If
        .TieneNecropsia = IIf(Me.chkSeRealizoNecropsia.Value, True, False)
        .HuboInfeccionIntraHospitalaria = False
        .NroReferenciaDestino = txtReferenciaD.Text
        'debb-21/06/2016 (inicio)
        '.referenciaOservicio
        '.referenciaOidDiagnostico
        .referenciaDservicio = PVcomboBoxDevuelveEleccion(cmbServicioReferenciaD)
        .referenciaDfextension = IIf(txtFextension.Text = sighEntidades.FECHA_VACIA_DMY, 0, txtFextension.Text)
        .referenciaDftramite = IIf(txtFtramite.Text = sighEntidades.FECHA_VACIA_DMY, 0, txtFtramite.Text)
        'debb-21/06/2016 (fin)
        If mo_Atenciones.IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS And wxParametro302 = "S" Then
                If .idSiaSis = 0 Or .SisCodigo = "" Then
                   mo_ReglasSISgalenhos.SisFiliacionesActualizarAfiliadoDesdeWEB Left(mo_Pacientes.nrodocumento, 8), "", "", _
                                        "", "", "", wxParametro323
                   mo_ReglasSISgalenhos.SisFiliacionesDevuelveKEY lnAfiliacionSIS4, lcSIScodigo, _
                                        mo_Pacientes.ApellidoPaterno, mo_Pacientes.ApellidoMaterno, _
                                        mo_Pacientes.PrimerNombre, mo_Pacientes.FechaNacimiento, _
                                        lcCodigoEstablecimientoAdscripcionSIS
                   .idSiaSis = lnAfiliacionSIS4
                   .SisCodigo = lcSIScodigo
                End If
        Else
                .idSiaSis = 0
                .FuaCodigoPrestacion = ""
                .SisCodigo = ""
        End If
        If ml_TipoServicio = sghEmergenciaConsultorios Then
           .idServicioDestino = Val(mo_cmbServicioDestino.BoundText)
        End If
    End With


 
    'debb2014b
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE DIAGNOSTICOS DE INGRESO
    '---------------------------------------------------------------------------------
    If SeLlamaDesdeCreacionDeCuentaEnHospitalizacion = False Then
        Me.ucDiagnosticosIngreso.idUsuario = ml_idUsuario
        ucDiagnosticosIngreso.TipoDiagnostico = sghHospitalizacionIngreso
        Me.ucDiagnosticosIngreso.CargarDiagnosticosAlObjetoDatos mo_Diagnosticos
    End If
    'debb2014b
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE DIAGNOSTICOS DE EGRESO
    '---------------------------------------------------------------------------------
    If SeLlamaDesdeCreacionDeCuentaEnHospitalizacion = False Then
        Me.ucDiagnosticosEgreso.idUsuario = ml_idUsuario
        Me.ucDiagnosticosEgreso.TipoDiagnostico = sghHospitalizacionEgreso
        Me.ucDiagnosticosEgreso.CargarDiagnosticosAlObjetoDatos mo_Diagnosticos
    Else
        Me.ucDiagnosticosEgreso.idUsuario = ml_idUsuario
        Me.ucDiagnosticosEgreso.TipoDiagnostico = sghHospitalizacionIngreso
        Me.ucDiagnosticosEgreso.CargarDiagnosticosAlObjetoDatos mo_Diagnosticos
    End If
    
    
    
    
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE DIAGNOSTICOS DE MORTALIDAD
    '---------------------------------------------------------------------------------
    If SeLlamaDesdeCreacionDeCuentaEnHospitalizacion = False Then
        Me.ucDiagnosticosMortalidad.idUsuario = ml_idUsuario
        Me.ucDiagnosticosMortalidad.TipoDiagnostico = sghHospitalizacionMortalidad
        Me.ucDiagnosticosMortalidad.CargarDiagnosticosAlObjetoDatos mo_Diagnosticos
    End If
    
    If SeLlamaDesdeCreacionDeCuentaEnHospitalizacion = False Then
        '---------------------------------------------------------------------------------
        '           CARGA DATOS DE NACIMIENTO
        '---------------------------------------------------------------------------------
        Me.ucNacimientoDetalle1.idUsuario = ml_idUsuario
        Me.ucNacimientoDetalle1.CargarNacimientosAlObjetoDatos mo_Nacimientos
        
        '---------------------------------------------------------------------------------
        '           CARGA DATOS DE DIAGNOSTICOS DE NACIMIENTO
        '---------------------------------------------------------------------------------
        Me.ucDiagnosticoNacimiento.idUsuario = ml_idUsuario
        Me.ucDiagnosticoNacimiento.TipoDiagnostico = sghHospitalizacionNacimiento
        Me.ucDiagnosticoNacimiento.CargarDiagnosticosAlObjetoDatos mo_Diagnosticos
    End If
    
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE DIAGNOSTICOS DE COMPLICACIONES
    '---------------------------------------------------------------------------------
    If SeLlamaDesdeCreacionDeCuentaEnHospitalizacion = False Then
        Me.ucDiagnosticoComplicaciones.idUsuario = ml_idUsuario
        Me.ucDiagnosticoComplicaciones.TipoDiagnostico = sghHospitalizacionComplicaciones
        Me.ucDiagnosticoComplicaciones.CargarDiagnosticosAlObjetoDatos mo_Diagnosticos
    End If
    
    Dim lcFechaEgreso As String, lcHoraEgreso As String
    lcFechaEgreso = Me.txtFechaEgreso.Text
    lcHoraEgreso = Me.txtHoraEgreso.Text
    If SeLlamaDesdeCreacionDeCuentaEnHospitalizacion = True Then
       lcFechaEgreso = sighEntidades.FECHA_VACIA_DMY
       lcHoraEgreso = sighEntidades.HORA_VACIA_HM
    End If
    Dim oDOOcupacion As New DOEstanciaHospitalaria
    oDOOcupacion.IdServicio = IIf(SeLlamaDesdeCreacionDeCuentaEnHospitalizacion = True, Val(mo_cmbServicioDestino.BoundText), mo_Atenciones.IdServicioIngreso)
    oDOOcupacion.IdMedicoOrdena = IIf(SeLlamaDesdeCreacionDeCuentaEnHospitalizacion = True, 0, mo_Atenciones.IdMedicoIngreso)
    oDOOcupacion.FechaOcupacion = IIf(SeLlamaDesdeCreacionDeCuentaEnHospitalizacion = True, mo_Atenciones.fechaEgreso, mo_Atenciones.FechaIngreso)
    oDOOcupacion.HoraOcupacion = IIf(SeLlamaDesdeCreacionDeCuentaEnHospitalizacion = True, mo_Atenciones.HoraEgreso, mo_Atenciones.HoraIngreso)
    oDOOcupacion.idCama = mo_Atenciones.IdCamaIngreso
    oDOOcupacion.IdUsuarioAuditoria = ml_idUsuario
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE TRANSFERENCIAS
    '---------------------------------------------------------------------------------
    If SeLlamaDesdeCreacionDeCuentaEnHospitalizacion = True Then
       Me.ucTransferenciasDetalle1.LimpiarDatos
    End If
    Me.ucTransferenciasDetalle1.idUsuario = ml_idUsuario
    Me.ucTransferenciasDetalle1.CargaTransferenciasAlObjetosDatos mo_OcupacionCamas, oDOOcupacion, lcFechaEgreso, _
            lcHoraEgreso, 1, 0, lnSecuenciaTransferencia, 0, mo_Atenciones.IdCamaEgreso
    
    If ml_TipoServicio = sghEmergenciaConsultorios Or ml_TipoServicio = sghEmergenciaObservacion Then
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE ATENCIONES DE EMERGENCIA
    '---------------------------------------------------------------------------------
        With mo_AtencionesEmergencia
            .comoLlego = Me.cmbComoLlego.ListIndex + 1
            .tipoAtencion = Me.cmbTipoAtencion.ListIndex + 1
            .idEstadoLlegada = IIf(Me.cmbEstadoLlegada.Text = "", 0, Me.cmbEstadoLlegada.ListIndex + 1)
        
        End With
    End If
End Sub
'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------




'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean
    If CalculaEstanciaParaPacienteConAltaMedica(mo_Atenciones) = True Then
        Dim oEpisodioClinico As EpisodioClinico
        oEpisodioClinico = EpisodioClinicoDevuelveDatos
        '
        ModificarDatos = mo_AdminAdmision.AdmisionHospModEGRESO(mo_CuentasAtencion, mo_Atenciones, _
                                                                mo_OcupacionCamas, _
                                                                mo_Diagnosticos, mo_Procedimientos, mo_Examenes, mo_Nacimientos, _
                                                                ml_TipoAccionAdmision, ldFechaEgresoMedicoAnterior, lbPacienteNN, _
                                                                mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, _
                                                                Me.Caption, _
                                                                oRsEstancia, lnIdNacimientoSeleccionado, _
                                                                mo_DoAtencionDatosAdicionales, _
                                                                mb_EsObservacionEmergencia, _
                                                                oEpisodioClinico, wxParametro511, _
                                                                mo_AtencionesEmergencia)
        ms_MensajeError = mo_AdminAdmision.MensajeError
        If ms_MensajeError = "" Then
            Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
            mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar mo_Atenciones.idCuentaAtencion, False, 0
            Set mo_ReglasFacturacion = Nothing
            '
            If wxParametro302 = "S" And mo_Atenciones.IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
               mo_ReglasSISgalenhos.SisFuaAtencionActualizaDatosDesdeHospEmegCE mo_Atenciones.idCuentaAtencion, _
                                                                      mo_Atenciones.idTipoServicio, mo_Atenciones.idAtencion, _
                                                                      mo_lnIdTablaLISTBARITEMS, ml_idUsuario
            End If
        End If
    End If
End Function


'debb-Jamo

Function EpisodioClinicoDevuelveDatos() As EpisodioClinico
        Dim oEpisodioClinico As EpisodioClinico
        oEpisodioClinico.idEpisodio = Me.UcEpisodioClinico1.idEpisodio
        oEpisodioClinico.lbCierreEpisodio = Me.UcEpisodioClinico1.lbCierreEpisodio
        oEpisodioClinico.lbNuevoEpisodio = Me.UcEpisodioClinico1.lbNuevoEpisodio
        EpisodioClinicoDevuelveDatos = oEpisodioClinico
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------




Sub CargarDatosAlosControles()
        Dim oConexion As New Connection
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighEntidades.CadenaConexion
        'CargaFormaPago
        
        Select Case ml_TipoServicio
        Case 2
            Set mo_cmbIdDestinoAtencion.RowSource = mo_AdminAdmision.TiposDestinoAtencionSeleccionarDestinosDeConsultorioEmergencia
        Case 3
            Set mo_cmbIdDestinoAtencion.RowSource = mo_AdminAdmision.TiposDestinoAtencionSeleccionarDestinosDeHospitalizacion(sghSoloPacHospitalizados)
        End Select
        
        '
        '1do:   CARGAR DATOS DE LA ATENCION
        CargarDatosDelaAtencion oConexion
        
        If mo_Atenciones.idAtencion = 0 Then
             mb_ExistenDatos = False
             Exit Sub
        End If
        '
        lbSeEnviaMensajeCelularAreaConv = False
        If mo_Atenciones.IdFormaPago <> 1 And mo_Atenciones.IdFormaPago <> 2 And mo_Atenciones.IdFormaPago <> 5 Then   'NO a: Particular,SIS, ParticularHospitalizado
           lbSeEnviaMensajeCelularAreaConv = True
        End If
        '
        If ml_TipoServicio = sghEmergenciaConsultorios Then
          fraSoloEme.Visible = True
          Dim rsTiposGravedad As New ADODB.Recordset
          Set rsTiposGravedad = mo_AdminServiciosComunes.TipoGravedadAtencionSeleccionarTodos()
          mo_cmbIdTipoGravedad.CargarComboBoxDesdeRecordset cmbIdTipoGravedad, rsTiposGravedad, "IdTipoGravedad", "Descripcion"
          mo_cmbIdTipoGravedad.BoundText = mo_Atenciones.IdTipoGravedad
        Else
          fraSoloEme.Visible = False
        End If
        '
        Set mo_CuentasAtencion = mo_AdminFacturacion.CuentasAtencionSeleccionarPorId(Me.idCuentaAtencion, oConexion)
        If mo_CuentasAtencion.idEstado <> 1 Then
            btnAceptar.Enabled = False
        End If
        ml_idCuentaAtencion = mo_CuentasAtencion.idCuentaAtencion
        ml_TipoFinanciamiento = mo_AdminServiciosComunes.FuentesFinanciamientoDevuelveDescripcion(mo_Atenciones.IdFuenteFinanciamiento)
        Me.Caption = "(HC: " & Trim(mo_Pacientes.NroHistoriaClinica) & " " & _
                     Trim(mo_Pacientes.ApellidoPaterno) & " " & Trim(mo_Pacientes.ApellidoMaterno) & _
                    " " & Trim(mo_Pacientes.PrimerNombre) & _
                    ")(Estado Cta: " & mo_ReglasFarmacia.DevuelveEstadoActualDeEstadoCuenta("idEstado=" & mo_CuentasAtencion.idEstado, oConexion) & _
                    ")(Edad: " & Trim(Str(mo_Atenciones.Edad)) & " " & sighEntidades.EdadDevuelveTipo(mo_Atenciones.idTipoEdad) & _
                    ")(T.F: " & ml_TipoFinanciamiento & _
                    ")(Gs: " & IIf(IsNull(mo_Pacientes.GrupoSanguineo), "", mo_Pacientes.GrupoSanguineo) & _
                    ")(Frh: " & IIf(IsNull(mo_Pacientes.FactorRh), "", mo_Pacientes.FactorRh) & ")"
        lblIdAtencion.Caption = mo_Atenciones.idAtencion
        
        '4to:   PARA VISUALIZAR LA UBICACION DEL PACIENTE AL DIA DE LA ATENCION
        mo_DoUbicacionPaciente.DireccionDomicilio = mo_DoAtencionDatosAdicionales.DireccionDomicilio
        
    
        'debb2014b
        '4to:   CARGAR DATOS DE LOS DIAGNOSTICOS INGRESO POR ATENCION
        Me.ucDiagnosticosIngreso.Inicializar
        Me.ucDiagnosticosIngreso.idAtencion = Me.idAtencion
        Me.ucDiagnosticosIngreso.TipoDiagnostico = sghHospitalizacionIngreso
        Me.ucDiagnosticosIngreso.CargarDatosDeDiagnosticos oConexion
        'debb2014b
        
        CambioFechaNacimiento Format(mo_Pacientes.FechaNacimiento, sighEntidades.DevuelveFechaSoloFormato_DMY), Format(mo_Pacientes.FechaNacimiento, sighEntidades.DevuelveHoraSoloFormato_HM)
        '5to:   CARGAR DATOS DE LOS DIAGNOSTICOS EGRESO POR ATENCION
        Me.ucDiagnosticosEgreso.idAtencion = Me.idAtencion
        Me.ucDiagnosticosEgreso.SexoPaciente = mo_Pacientes.idTipoSexo
        Me.ucDiagnosticosEgreso.CargarDatosDeDiagnosticos oConexion
        
        '6to:   CARGAR DATOS DE LOS DIAGNOSTICOS MORTALIDAD POR ATENCION
        Me.ucDiagnosticosMortalidad.idAtencion = Me.idAtencion
        Me.ucDiagnosticosMortalidad.SexoPaciente = mo_Pacientes.idTipoSexo
        Me.ucDiagnosticosMortalidad.CargarDatosDeDiagnosticos oConexion
        
        '7to:   CARGAR DATOS DE LOS DIAGNOSTICOS NACIMIENTO POR ATENCION
        Me.ucDiagnosticoNacimiento.idAtencion = Me.idAtencion
        Me.ucDiagnosticoNacimiento.SexoPaciente = mo_Pacientes.idTipoSexo
        Me.ucDiagnosticoNacimiento.CargarDatosDeDiagnosticos oConexion
        
        '8to:   CARGAR DATOS DE LOS DIAGNOSTICOS COMPLICACIONES POR ATENCION
        Me.ucDiagnosticoComplicaciones.idAtencion = Me.idAtencion
        Me.ucDiagnosticoComplicaciones.SexoPaciente = mo_Pacientes.idTipoSexo
        Me.ucDiagnosticoComplicaciones.CargarDatosDeDiagnosticos oConexion
        
        '11to:    CARGAR DATOS DE OCUPACION DE EGRESOS
        lnSecuenciaTransferencia = 0
        Me.ucTransferenciasDetalle1.FechaIngreso = CDate(Format(mo_Atenciones.FechaIngreso, sighEntidades.DevuelveFechaSoloFormato_DMY) & " " & mo_Atenciones.HoraIngreso)
        Me.ucTransferenciasDetalle1.idAtencion = Me.idAtencion
        Me.ucTransferenciasDetalle1.idCuentaAtencion = mo_Atenciones.idCuentaAtencion
        Me.ucTransferenciasDetalle1.CargarDatosDeTransferencias oConexion
        CompletarDatosDeEgreso mo_Atenciones.IdServicioEgreso, mo_Atenciones.IdCamaEgreso
        
        'If ml_TipoServicio = sghHospitalizacion Then
            '12to:    CARGAR DATOS DE OCUPACION DE CAMAS
            Me.ucNacimientoDetalle1.idAtencion = Me.idAtencion
            Me.ucNacimientoDetalle1.CargarDatosDeNacimientos oConexion
        'End If
        
        '13to:    CARGAR DATOS DE ATENCION DE EMREGENCIA
        If ml_TipoServicio = sghEmergenciaConsultorios Or ml_TipoServicio = sghEmergenciaObservacion Then
            CargarDatosDeLaAtencionDeEmergencia oConexion
        End If
        
        '14avo:    CARGAR FECHA DE EGRESO MEDICO, si lo tuviese, para "MODIFICAR"--> con el fin de generar CONSUMO POR DIAS DE ESTANCIA
        ldFechaEgresoMedicoAnterior = CDate(IIf(txtFechaEgreso.Text = sighEntidades.FECHA_VACIA_DMY, "01/01/1900", txtFechaEgreso.Text))
        
        '
        If wxParametro215 = "0" Then 'inhabilita Fecha/hora Administrativa
           txtHoraEgresoAdm.Enabled = False
           txtFechaEgresoAdm.Enabled = False
        End If
        'Ya tuvo movimientos(Farmacia/servicios), no podrá cambiar de plan
        '
        '
        '
        '
        '
        If mi_Opcion <> sghAgregar Then
           Me.UcEpisodioClinico1.idPaciente = mo_Atenciones.idPaciente
           Me.UcEpisodioClinico1.idAtencion = mo_Atenciones.idAtencion
           Me.UcEpisodioClinico1.Inicializar
           Me.UcEpisodioClinico1.Limpiar
           Me.UcEpisodioClinico1.CargaEpisodiosHistoricos
           If Me.txtFechaEgreso <> sighEntidades.FECHA_VACIA_DMY Then
              Me.UcEpisodioClinico1.CargarDatosAlosControles oConexion
           End If
        End If
        '
        oConexion.Close
        Set oConexion = Nothing
        '
        HaceVisibleOnoBotonFUA
        If mo_Atenciones.IdDestinoAtencion > 0 Then
          mo_cmbIdDestinoAtencion.BoundText = mo_Atenciones.IdDestinoAtencion
        End If
        
End Sub



Sub CargarDatosDeLasAtencionesPadres()
    
    '---------------------------------------------------------------------------------
    'CARGAR DATOS DE LA ATENCION PADRE SI ES QUE HUBIERA
    '---------------------------------------------------------------------------------
    Dim oConexion As New Connection
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Select Case ml_TipoAccionAdmision

    Case sghIngresarUnAlojamientoConjunto
    
        Set mo_AtencionPadre = mo_AdminAdmision.AtencionesSeleccionarPorId(ml_IdAtencionPadre, oConexion)
        Set mo_CuentasAtencion = mo_AdminFacturacion.CuentasAtencionSeleccionarPorId(mo_AtencionPadre.idCuentaAtencion, oConexion)
        
    Case sghEnviarAObservacion
        
        Set mo_AtencionPadre = mo_AdminAdmision.AtencionesSeleccionarPorId(ml_IdAtencionPadre, oConexion)
        mo_AtencionPadre.fechaEgreso = Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY)
        mo_AtencionPadre.HoraEgreso = Format(Now, sighEntidades.DevuelveHoraSoloFormato_HM)
        mo_AtencionPadre.IdTipoAlta = 5         'Paciente trasladado a observacion de emergencia
        mo_AtencionPadre.IdCondicionAlta = 3    'Paciente inalterado
        'Se supone que a observacion siempre llegan desde el consultorio
        mo_AtencionPadre.IdDestinoAtencion = 22
        
        Set mo_CuentasAtencion = mo_AdminFacturacion.CuentasAtencionSeleccionarPorId(mo_AtencionPadre.idCuentaAtencion, oConexion)
        
    Case sghTrasladarAHospitalizacion
        
        Set mo_AtencionPadre = mo_AdminAdmision.AtencionesSeleccionarPorId(ml_IdAtencionPadre, oConexion)
        mo_AtencionPadre.fechaEgreso = Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY)
        mo_AtencionPadre.HoraEgreso = Format(Now, sighEntidades.DevuelveHoraSoloFormato_HM)
        mo_AtencionPadre.IdTipoAlta = 6         'Paciente trasladado a hospitalizaciòn
        mo_AtencionPadre.IdCondicionAlta = 3    'Paciente inalterado
        
        'Si viene a hospitalizarse desde de observacion de emergencia
        If mo_AtencionPadre.idTipoServicio = sghEmergenciaConsultorios Then
            mo_AtencionPadre.IdDestinoAtencion = 21
        End If
        
        'Si viene a hospitalizarse desde consultorio de emergencia
        If mo_AtencionPadre.idTipoServicio = sghEmergenciaObservacion Then
            mo_AtencionPadre.IdDestinoAtencion = 41
        End If
        
        Set mo_CuentasAtencion = mo_AdminFacturacion.CuentasAtencionSeleccionarPorId(mo_AtencionPadre.idCuentaAtencion, oConexion)
        
    End Select
    oConexion.Close
    Set oConexion = Nothing
End Sub

Sub CompletarDatosDeEgreso(lIdServicioEgreso As Long, lIdCamaEgreso As Long)
Dim oDOCama As New DOCama
Dim oDoServicio As New doServicio
Dim oConexion As New Connection
        oConexion.Open sighEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        Set oDOCama = mo_AdminHoteleria.CamasSeleccionarPorId(lIdCamaEgreso, oConexion)
        Me.txtNroCamaEgreso = oDOCama.Codigo
        Me.txtNroCamaEgreso.Tag = oDOCama.idCama
        mb_EsObservacionEmergencia = False
        Set oDoServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(lIdServicioEgreso, oConexion)
        If Not oDoServicio Is Nothing Then
            Me.txtIdServicioEgreso.Tag = oDoServicio.IdServicio
            Me.txtIdServicioEgreso.Text = oDoServicio.Codigo
            Me.lblNombreServicioEgreso = oDoServicio.nombre
            Me.lblNombreServicioEgreso.Tag = oDoServicio.IdEspecialidad
            mb_EsObservacionEmergencia = oDoServicio.EsObservacionEmergencia
            If mb_EsObservacionEmergencia = True Then    '09/08/2011
               Me.txtNroCamaEgreso.Visible = True
               Me.lblNroCamaEgreso.Visible = True
            End If
        Else
            Me.txtIdServicioEgreso.Tag = ""
            Me.lblNombreServicioEgreso = ""
            Me.lblNombreServicioEgreso.Tag = ""
        End If
        
        oConexion.Close
        Set oConexion = Nothing
        Set oDOCama = Nothing
        Set oDoServicio = Nothing
End Sub
Sub CargarDatosDeLaAtencionDeEmergencia(oConexion As Connection)

        Me.IdAtencionEmergencia = mo_AdminAdmision.AtencionesEmergenciaSeleccionarIdPorIdAtencion(ml_idAtencion, oConexion)
        If Me.IdAtencionEmergencia = 0 Then
             Exit Sub
        End If
        
        Set mo_AtencionesEmergencia = mo_AdminAdmision.AtencionesEmergenciaSeleccionarPorId(Me.IdAtencionEmergencia, oConexion)
        If mo_AdminAdmision.MensajeError <> "" Then
             MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminServiciosComunes.MensajeError, vbInformation, Me.Caption"
             Exit Sub
        End If
        
        If Not mo_AtencionesEmergencia Is Nothing Then
            With mo_AtencionesEmergencia
                If .IdCausaExternaMorbilidad = 0 Then
                   MsgBox "Tiene que elegir la CAUSA EXTERNA DE MORBILIDAD antes de registrar el ALTA" & _
                          Chr(13) & "opción Modificar->ficha 2.3 CAUSAS EXTERNAS DE MORBILIDAD", vbInformation, ""
                   btnAceptar.Visible = False
                End If
                If .comoLlego > 0 Then
                   cmbComoLlego.ListIndex = .comoLlego - 1
                End If
                If .tipoAtencion > 0 Then
                   Me.cmbTipoAtencion.ListIndex = .tipoAtencion - 1
                End If
                If .idEstadoLlegada > 0 Then
                   Me.cmbEstadoLlegada.ListIndex = .idEstadoLlegada - 1
                End If
                
            End With
        End If
                

End Sub
Sub CargarDatosDelaAtencion(oConexion As Connection)
Dim oDoMedico As New DOMedico
Dim oDOEmpleado As New dOEmpleado
Dim oDOEspecialidades As New Collection
Dim oDoServicio As New doServicio
Dim lcEstadoAtencion As String
        If lbCargaAlaVezCitaPacienteAtencionDA = False Then
           Set mo_Atenciones = mo_AdminAdmision.AtencionesSeleccionarPorId(Me.idAtencion, oConexion)
           If mo_Atenciones.idAtencion = 0 Then
                'El registro ha sido eliminado, pero no se hizo el refresh
                 mb_ExistenDatos = False
                 Exit Sub
           End If
        Else
           mo_Atenciones.idAtencion = Me.idAtencion
           mb_ExistenDatos = mo_AdminAdmision.AtencionesPacientesCitasDatosadicionalesSeleccionarPorId(mo_Pacientes, _
                                                mo_Atenciones, mo_DoAtencionDatosAdicionales, _
                                                oConexion, mo_CuentasAtencion, False)
           If mo_Atenciones.idAtencion = 0 Then
                'El registro ha sido eliminado, pero no se hizo el refresh
                 mb_ExistenDatos = False
                 Exit Sub
           End If
        End If
        If mo_AdminAdmision.MensajeError <> "" Then
             MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminServiciosComunes.MensajeError, vbInformation, Me.Caption"
             mb_ExistenDatos = False
             Exit Sub
        End If
        
        If Not mo_Atenciones Is Nothing Then
           '
           Me.ucNacimientoDetalle1.Visible = True
           Frame5.Visible = True
           ucDiagnosticoNacimiento.Visible = True
           If mo_AdminServiciosHosp.ServiciosPuedeUsarFichaNacimientos(IIf(mo_Atenciones.IdServicioEgreso = 0, mo_Atenciones.IdServicioIngreso, mo_Atenciones.IdServicioEgreso), oConexion) = False Then
               Me.ucNacimientoDetalle1.Visible = False
               Frame5.Visible = False
               ucDiagnosticoNacimiento.Visible = False
           End If
           '
           With mo_Atenciones
                Me.idAtencion = .idAtencion
                Me.idPaciente = .idPaciente
                
                mo_cmbIdDestinoAtencion.BoundText = .IdDestinoAtencion
                
                Me.txtIdMedicoEgreso.Tag = .IdMedicoEgreso
                
                Me.IdEspecialidad = .IdEspecialidadMedico
                
                
                
                
                Me.txtHoraEgreso.Text = IIf(.HoraEgreso = "", sighEntidades.HORA_VACIA_HM, .HoraEgreso)
                Me.txtFechaEgreso.Text = IIf(.fechaEgreso = 0, sighEntidades.FECHA_VACIA_DMY, .fechaEgreso)
                
                'Se guarda en estas variables para validar si el paciente ya esta de alta o no
                Me.txtHoraEgreso.Tag = Me.txtHoraEgreso.Text
                Me.txtFechaEgreso.Tag = Me.txtFechaEgreso.Text
                
                
                
                
                If mo_AdminProgramacion.MedicosSeleccionarPorId(.IdMedicoEgreso, oDoMedico, oDOEmpleado, oDOEspecialidades, oConexion) Then
                    Me.txtIdMedicoEgreso = oDOEmpleado.CodigoPlanilla
                    Me.lblNombreMedicoEgreso = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
                Else
                    Me.lblNombreMedicoEgreso = ""
                    txtIdMedicoEgreso.Text = ""
                End If
                
                If mo_AdminProgramacion.MedicosSeleccionarPorId(mo_DoAtencionDatosAdicionales.IdMedicoRespNacimiento, oDoMedico, oDOEmpleado, oDOEspecialidades, oConexion) Then
                    Me.txtIdMedicoNacimiento = oDOEmpleado.CodigoPlanilla
                    Me.lblNombreMedicoNacimiento = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
                Else
                    Me.lblNombreMedicoNacimiento = ""
                End If
                
                
                Set oDoServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(.IdServicioIngreso, oConexion)
                If Not oDoServicio Is Nothing Then
                    lcElServicioUsaGalenHos = IIf(oDoServicio.UsaGalenHos = True, "S", "N")
                    Me.IdEspecialidad = oDoServicio.IdEspecialidad
                Else
                End If
                If .FechaEgresoAdministrativo = 0 Then
                   Me.txtFechaEgresoAdm = sighEntidades.FECHA_VACIA_DMY
                Else
                   Me.txtFechaEgresoAdm.Text = Format(.FechaEgresoAdministrativo, sighEntidades.DevuelveFechaSoloFormato_DMY)
                End If
                If sighEntidades.EsHora(.HoraEgresoAdministrativo) Then
                   Me.txtHoraEgresoAdm = IIf(.HoraEgresoAdministrativo = "", sighEntidades.HORA_VACIA_HM, Format(.HoraEgresoAdministrativo, sighEntidades.DevuelveHoraSoloFormato_HM))
                Else
                   Me.txtHoraEgresoAdm = sighEntidades.HORA_VACIA_HM
                End If
                mo_cmbCondicionAlta.BoundText = .IdCondicionAlta
                mo_cmbTipoAlta.BoundText = .IdTipoAlta
                
                'Cama de ingreso
                                         
                
                'WCG comentado por facturacion
                Me.ucNacimientoDetalle1.FechaIngreso = IIf(.FechaIngreso = 0, sighEntidades.FECHA_VACIA_DMY, .FechaIngreso)
                
                
                'Me.chkSeRealizoNecropsia.Value = IIf(.TieneNecropsia, 1, 0)
                
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
                
                
                lnIdPlanAnterior = .IdFuenteFinanciamiento
                lnIdTipoFinanciamientoAnterior = .IdFormaPago
                ml_ldFechaIngreso = .FechaIngreso
                ml_lcHoraIngreso = .HoraIngreso
                ml_TipoServicio = .idTipoServicio
                ml_IdFuenteFinanciamiento = .IdFuenteFinanciamiento
                ml_IdViasAdmision = .IdOrigenAtencion
                'debb-03/07/2016
                If .IdTipoGravedad = 0 And ml_TipoServicio = sghEmergenciaConsultorios Then
                   MsgBox "Tiene que elegir la LA GRAVEDAD antes de registrar el ALTA" & _
                          Chr(13) & "opción Modificar->ficha 2.1 INGRESOS", vbInformation, ""
                   btnAceptar.Visible = False
                End If
                If ml_TipoServicio = sghHospitalizacion Then
                   If .IdMedicoIngreso = 0 Then
                        MsgBox "Tiene que elegir el MEDICO DE INGRESO antes de registrar el ALTA" & _
                               Chr(13) & "opción Modificar->ficha 2.1 INGRESOS", vbInformation, ""
                        btnAceptar.Visible = False
                   End If
                   If .IdCamaIngreso = 0 Then
                        MsgBox "Tiene que elegir LA CAMA DE INGRESO antes de registrar el ALTA" & _
                               Chr(13) & "opción Modificar->ficha 2.1 INGRESOS", vbInformation, ""
                        btnAceptar.Visible = False
                   End If
                End If
                '
                mb_ExistenDatos = True
           End With
           Me.ucNacimientoDetalle1.FechaIngreso = CDate(Format(ml_ldFechaIngreso, sighEntidades.DevuelveFechaSoloFormato_DMY) & " " & ml_lcHoraIngreso)
           '
           If lbCargaAlaVezCitaPacienteAtencionDA = False Then
              Set mo_DoAtencionDatosAdicionales = mo_AdminAdmision.AtencionesDatosAdicionalesSeleccionarPorId(Me.idAtencion, oConexion)
           End If
           With mo_DoAtencionDatosAdicionales
                Me.txtIdMedicoNacimiento.Tag = .IdMedicoRespNacimiento
                mo_cmbIdTipoReferenciaDestino.BoundText = .IdTipoReferenciaDestino
                CompletarDatosDelEstablecimientoEnElLoad .idEstablecimientoDestino, .IdEstablecimientoNoMinsaDestino, txtIdEstablecimientoDestino, lblNombreDestinoReferencia, .IdTipoReferenciaDestino
                txtReferenciaD.Text = IIf(IsNull(.NroReferenciaDestino), "", .NroReferenciaDestino)
                Me.chkSeRealizoNecropsia.Value = IIf(.TieneNecropsia, 1, 0)
                
                lnAfiliacionSIS4 = .idSiaSis
                lcSIScodigo = .SisCodigo
                'debb-21/06/2016 (inicio)
                PVcomboBoxUbicaPosicion .referenciaDservicio, cmbServicioReferenciaD
                txtFextension.Text = IIf(.referenciaDfextension = 0, sighEntidades.FECHA_VACIA_DMY, .referenciaDfextension)
                txtFtramite.Text = IIf(.referenciaDftramite = 0, sighEntidades.FECHA_VACIA_DMY, .referenciaDftramite)
                'debb-21/06/2016 (fin)
                If ml_TipoServicio = sghEmergenciaConsultorios Then
                   mo_cmbServicioDestino.BoundText = .idServicioDestino
                End If
                
           End With
           'ESTOS DATOS SE UTILIZARAN MAS ADELANTE PARA ACTUALIZAR LA UBICACION DE PACIENTE
           Dim oPacientesTmp As New SIGHComun.doPaciente
           If lbCargaAlaVezCitaPacienteAtencionDA = False Then
              Set oPacientesTmp = mo_AdminAdmision.PacientesSeleccionarPorId(ml_IdPaciente, oConexion)
           Else
              Set oPacientesTmp = mo_Pacientes
           End If
           If mo_AdminAdmision.MensajeError <> "" Then
                 MsgBox "No se pudo obtener los datos" + Chr(13) + mo_AdminAdmision.MensajeError, vbInformation, "Datos de paciente"
                 Exit Sub
           End If
           If Not oPacientesTmp Is Nothing Then
                lcHistoriaYpaciente = "(" & Trim(Str(oPacientesTmp.NroHistoriaClinica)) & ") " & Trim(oPacientesTmp.ApellidoPaterno) & " " & Trim(oPacientesTmp.ApellidoMaterno) & " " & Trim(oPacientesTmp.PrimerNombre)
'                Me.Caption = Trim(Me.Caption) & "  (HC: " & Trim(oPacientesTmp.NroHistoriaClinica) & " " & _
'                             Trim(oPacientesTmp.ApellidoPaterno) & " " & Trim(oPacientesTmp.ApellidoMaterno) & _
'                            " " & Trim(oPacientesTmp.PrimerNombre) & ") (Estado: " & lcEstadoAtencion & _
'                            ")(Gs: " & IIf(IsNull(oPacientesTmp.GrupoSanguineo), "", oPacientesTmp.GrupoSanguineo) & _
'                            ", Frh: " & IIf(IsNull(oPacientesTmp.FactorRh), "", oPacientesTmp.FactorRh) & ")"
                With oPacientesTmp
                    mo_DoUbicacionPaciente.IdPaisDomicilio = .IdPaisDomicilio
                    mo_DoUbicacionPaciente.IdCentroPobladoDomicilio = .IdCentroPobladoDomicilio
                    
                    mo_DoUbicacionPaciente.IdPaisProcedencia = .IdPaisProcedencia
                    mo_DoUbicacionPaciente.IdCentroPobladoProcedencia = .IdCentroPobladoProcedencia
                    
                    mo_DoUbicacionPaciente.DireccionDomicilio = .DireccionDomicilio
                End With
           End If
           Set oPacientesTmp = Nothing
           '
       Else
           mb_ExistenDatos = False
           Exit Sub
       End If
       
        Set oDoMedico = Nothing
        Set oDOEmpleado = Nothing
        Set oDOEspecialidades = Nothing
        Set oDoServicio = Nothing

End Sub




'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla CuentasAtencion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()

           'LIMPIAR DATOS DE LA CUENTA DE ATENCION
           Me.idCuentaAtencion = 0
           Me.idAtencion = 0
           
           'LIMPIAR DATOS DE LA ATENCION
            
           Me.idAtencion = 0
           Me.IdAtencionEmergencia = 0
           
           
End Sub
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'*****************************************************************************************
'                               EVENTOS DE LA ATENCION
'*****************************************************************************************
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------





















































Sub CompletarDatosDeDiagnostico(txtCodigoDx As TextBox, lblDescripcionDx As TextBox)
Dim oBusqueda As New SIGHNegocios.BuscaDiagnosticos
Dim oDODiagnostico As DODiagnostico

    oBusqueda.MostrarFormulario
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDODiagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorId(oBusqueda.idRegistroSeleccionado)
        If Not oDODiagnostico Is Nothing Then
            txtCodigoDx.Text = oDODiagnostico.CodigoCIE2004
            txtCodigoDx.Tag = oDODiagnostico.idDiagnostico
            lblDescripcionDx = oDODiagnostico.descripcion
        Else
            txtCodigoDx.Text = ""
            txtCodigoDx.Tag = ""
            lblDescripcionDx = ""
        End If
    End If
    Set oBusqueda = Nothing
End Sub
Sub CompletarDatosDeDiagnosticoEnELLostFocus(txtCodigoDx As TextBox, lblDescripcionDx As TextBox)
    
    txtCodigoDx.Text = UCase(txtCodigoDx.Text)

   If txtCodigoDx.Text <> "" Then
    Dim oDODiagnostico As DODiagnostico
        Set oDODiagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorCodigoCIE2004(txtCodigoDx.Text, True)
        If Not oDODiagnostico Is Nothing Then
            txtCodigoDx.Tag = oDODiagnostico.idDiagnostico
            lblDescripcionDx = oDODiagnostico.descripcion
        Else
            txtCodigoDx.Tag = ""
            lblDescripcionDx = ""
        End If
   End If

End Sub
Sub CompletarDatosDeMedico(txtMedico As TextBox, lblNombreMedico As TextBox, lIdEspecialidad As Long, lcFiltraMedico As String, ldFechaProgramada As Date, lcHoraProgramada As String, lnIdTipoServicio As Long)
'Dim oBusqueda As New MedicosBusqueda
Dim oBusqueda As New SIGHNegocios.BuscaMedicos
Dim oDoMedico As New DOMedico
Dim oDOEmpleado As New dOEmpleado
Dim oDOEspecialidades As New Collection
Dim oConexion As New Connection
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
     
    oBusqueda.IdEspecialidad = lIdEspecialidad
    oBusqueda.FechaProgramada = ldFechaProgramada
    oBusqueda.HoraProgramada = lcHoraProgramada
    oBusqueda.idTipoServicio = lnIdTipoServicio
    If mi_Opcion = sghAgregar Then
        oBusqueda.NombreMedico = lcFiltraMedico
    End If
    'oBusqueda.Show 1
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        If mo_AdminProgramacion.MedicosSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oDoMedico, oDOEmpleado, oDOEspecialidades, oConexion) Then
            txtMedico.Text = oDOEmpleado.CodigoPlanilla
            txtMedico.Tag = oDoMedico.idMedico
            lblNombreMedico = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oBusqueda = Nothing
    Set oDoMedico = Nothing
    Set oDOEmpleado = Nothing
    Set oDOEspecialidades = Nothing
End Sub

Sub CompletarDatosDeMedicoEgreso(txtMedico As TextBox, lblNombreMedico As TextBox, lIdEspecialidad As Long, lcFiltraMedico As String, ldFechaProgramada As Date, lcHoraProgramada As String, lnIdTipoServicio As Long)
'Dim oBusqueda As New MedicosBusqueda
Dim oBusqueda As New SIGHNegocios.BuscaMedicos
Dim oDoMedico As New DOMedico
Dim oDOEmpleado As New dOEmpleado
Dim oDOEspecialidades As New Collection
Dim oConexion As New Connection
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient

'    txtMedico.Tag = "" 'Actualizado 10102014
'    lblNombreMedico.Text = "" 'Actualizado 10102014

    oBusqueda.IdEspecialidad = lIdEspecialidad
    oBusqueda.NombreMedico = lcFiltraMedico
    oBusqueda.FechaProgramada = ldFechaProgramada
    oBusqueda.HoraProgramada = lcHoraProgramada
    oBusqueda.idTipoServicio = lnIdTipoServicio
    oBusqueda.NoMuestraInactivos = True
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        If mo_AdminProgramacion.MedicosSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oDoMedico, oDOEmpleado, oDOEspecialidades, oConexion) Then
            txtMedico.Text = oDOEmpleado.CodigoPlanilla
            txtMedico.Tag = oDoMedico.idMedico
            lblNombreMedico = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oBusqueda = Nothing
    Set oDoMedico = Nothing
    Set oDOEmpleado = Nothing
    Set oDOEspecialidades = Nothing
End Sub


Sub CompletarDatosDeMedicoEnElLostFocus(txtMedico As TextBox, lblNombreMedico As TextBox)
Dim oMedicosEspecialidad As New Collection
    txtMedico = Trim(txtMedico)
    If txtMedico <> "" Then
        Dim oDOEmpleado As New dOEmpleado
        Dim oDoMedico As New DOMedico
        If mo_AdminProgramacion.MedicosSeleccionarPorCodigo1(CStr(txtMedico), oDoMedico, oDOEmpleado, oMedicosEspecialidad) Then
            txtMedico.Tag = oDoMedico.idMedico
            Set oDOEmpleado = mo_AdminServiciosComunes.EmpleadosSeleccionarPorId(oDoMedico.IdEmpleado)
            lblNombreMedico = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        Else
            txtMedico.Tag = ""
            lblNombreMedico = ""
        End If
        Set oDOEmpleado = Nothing
        Set oDoMedico = Nothing
    Else
    End If
    Set oMedicosEspecialidad = Nothing
End Sub

Sub CompletarDatosDeServicio(txtIdServicio As TextBox, lblDescripcionServicio As TextBox, lcFiltraServicio As String)
Dim oBusqueda As New SIGHNegocios.BuscaServicioHosp
Dim oDoServicio As New doServicio
Dim oConexion As New Connection
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    oBusqueda.idTipoServicio = ml_TipoServicio
    oBusqueda.HabilitarTipoServicio = False
    If mi_Opcion = sghAgregar Then
        oBusqueda.NombreServicio = lcFiltraServicio
    End If
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDoServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oConexion)
        If Not oDoServicio Is Nothing Then
            If ml_TipoServicio = oDoServicio.idTipoServicio Then
                lcElServicioUsaGalenHos = IIf(oDoServicio.UsaGalenHos = True, "S", "N")
                txtIdServicio.Text = oDoServicio.Codigo
                txtIdServicio.Tag = oDoServicio.IdServicio
                lblDescripcionServicio = oDoServicio.nombre
                lblDescripcionServicio.Tag = oDoServicio.IdEspecialidad
                If ml_TipoServicio = sghEmergenciaConsultorios Then   '09/08/2011
                    mb_EsObservacionEmergencia = False
                    If oDoServicio.EsObservacionEmergencia = True Then
                        mb_EsObservacionEmergencia = True
                    End If
                End If
            Else
                MsgBox "El servicio seleccionado no pertenece a emergencia", vbInformation, Me.Caption
                txtIdServicio.Text = ""
                txtIdServicio.Tag = ""
                lblDescripcionServicio = ""
                lblDescripcionServicio.Tag = ""
                lcElServicioUsaGalenHos = "N"
            End If
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oBusqueda = Nothing
    Set oDoServicio = Nothing
    
End Sub



Sub CompletarDatosDeServicioEnElLostFocus(txtIdServicio As TextBox, lblDescripcionServicio As TextBox)
    
    txtIdServicio.Text = UCase(txtIdServicio.Text)
    If txtIdServicio.Text <> "" Then
        Dim oDoServicio As doServicio
        Set oDoServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorCodigo(txtIdServicio.Text)
        If Not oDoServicio Is Nothing Then
            If ml_TipoServicio = oDoServicio.idTipoServicio Then
                txtIdServicio.Tag = oDoServicio.IdServicio
                lblDescripcionServicio = oDoServicio.nombre
                lblDescripcionServicio.Tag = oDoServicio.IdEspecialidad
            Else
                MsgBox "El servicio ingresado no pertenece es de emergencia", vbInformation, Me.Caption
                txtIdServicio.Tag = ""
                lblDescripcionServicio = ""
                lblDescripcionServicio.Tag = ""
            End If
        Else
            txtIdServicio.Tag = ""
            lblDescripcionServicio = ""
            lblDescripcionServicio.Tag = ""
        End If
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
                Set oDoEstablecimiento = Nothing
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
                Set oDOEstablecimientoNoMinsa = Nothing
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
        Set oDoEstablecimiento = Nothing
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
         Set oDOEstablecimientoNoMinsa = Nothing
    End If

End Sub


Sub SePrecionoF2EnDx(lnKeyCode As Integer)
    On Error Resume Next
    Select Case lnKeyCode
    Case vbKeyTab
       If btnAceptar.Enabled Then btnAceptar.SetFocus
    Case vbKeyF2
       If btnAceptar.Enabled Then btnAceptar_Click
    Case Else
       AdministrarKeyPreview lnKeyCode
    End Select

End Sub









Private Sub txtReferenciaD_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdEstablecimientoDestino
End Sub

Private Sub ucDiagnosticosEgreso_SeIngresoDx(lcDx As String, SeElimino As Boolean)
    ActualizaPrioridadDeEmergencia lcDx, SeElimino
End Sub

Sub ActualizaPrioridadDeEmergencia(lcDx As String, SeElimino As Boolean)
    On Error Resume Next
    If ml_TipoServicio = sghEmergenciaConsultorios And wxParametro559 = "S" Then
        If lcDx <> "" Then
           Dim rsTiposGravedad As New ADODB.Recordset
           Dim lcPrioridad As String
           If SeElimino = True Then
              'lnIdTipoGravedadEgreso = 0
           Else
              lcPrioridad = mo_AdminServiciosComunes.DevuelvePrioridadEmergencia(lcDx)
              If lcPrioridad <> "" Then
                 Set rsTiposGravedad = mo_AdminServiciosComunes.TipoGravedadAtencionSeleccionarTodos()
                 rsTiposGravedad.MoveFirst
                 Do While Not rsTiposGravedad.EOF
                    If InStr(rsTiposGravedad!descripcion, "sem=" & lcPrioridad) > 0 Then
                       lnIdTipoGravedadEgreso = rsTiposGravedad!IdTipoGravedad
                       mo_cmbIdTipoGravedad.BoundText = Trim(Str(lnIdTipoGravedadEgreso))
                       Exit Do
                    End If
                    rsTiposGravedad.MoveNext
                 Loop
              End If
           End If
           Set rsTiposGravedad = Nothing
        End If
    End If
End Sub

Private Sub ucDiagnosticosEgreso_SePresionoTeclaEspecial(KeyCode As Integer)
    SePrecionoF2EnDx (KeyCode)
End Sub




Private Sub ucDiagnosticosMortalidad_SeIngresoDx(lcDx As String, SeElimino As Boolean)
    ActualizaPrioridadDeEmergencia lcDx, SeElimino
End Sub

Private Sub ucDiagnosticosMortalidad_SePresionoTeclaEspecial(KeyCode As Integer)
    SePrecionoF2EnDx (KeyCode)
End Sub


Private Sub ucNacimientoDetalle1_SePresionoTeclaEspecial(KeyCode As Integer)
    Select Case KeyCode
    Case 1000  'Se pulso boton AGREGAR
        Dim oEdad As Edad
        oEdad = sighEntidades.CalcularEdad(ucNacimientoDetalle1.FechaNacimiento, _
                         CDate(Format(ml_ldFechaIngreso, sighEntidades.DevuelveFechaSoloFormato_DMY) & " " & ml_lcHoraIngreso))
        Me.ucDiagnosticoNacimiento.EdadPaciente = sighEntidades.EdadEnDias(oEdad)
        Me.ucDiagnosticoNacimiento.SexoPaciente = ucNacimientoDetalle1.idTipoSexo
    End Select
End Sub



Sub CambioFechaNacimiento(sFechaNacimiento As String, sHoraNacimiento As String)
    On Error Resume Next
    Dim oEdad As Edad
    oEdad = sighEntidades.CalcularEdad(CDate(sFechaNacimiento & " " & sHoraNacimiento), CDate(Format(ml_ldFechaIngreso, sighEntidades.DevuelveFechaSoloFormato_DMY) & " " & ml_lcHoraIngreso))
    Me.ucDiagnosticoComplicaciones.EdadPaciente = EdadEnDias(oEdad)
    Me.ucDiagnosticoNacimiento.EdadPaciente = EdadEnDias(oEdad)
    Me.ucDiagnosticosEgreso.EdadPaciente = EdadEnDias(oEdad)
    Me.ucDiagnosticosMortalidad.EdadPaciente = EdadEnDias(oEdad)
    
End Sub




























































































Sub LimpiarVariablesDeMemoria()

End Sub

Function CalculaEstanciaParaPacienteConAltaMedica(oDOAtencion As DOAtencion) As Boolean
        Dim lnDiasEstancia As Integer: Dim lnHorasEstancia As Integer
        Dim lcCodServicio As String: Dim lcHoraEstanciaMax  As String
        Dim lbEliminaConsultaEmerg As Boolean, lnIdServicio As Long
        Dim lcCodigoNombre As String, lnPrecioUnitario As Double
        Dim lnIdTipoFinanciamiento As Long, lnIdProducto As Long
        Dim lnTotalDiasEstancia As Long, lnTotalPagarEstancia As Double
        Dim oBuscaDiasPaciente As New SIGHDatos.Parametros
        Dim oRsTmp As New Recordset
        Dim oDoCatalogoServicioHosp As New DOFinanciamientoCatalogoServ
        Dim oGenerarRecordsetProductos As New SighFacturacion.dllFactUcEstadoCuenta
        Dim oConexion As New Connection
        Dim lnPrecioSIS As Double, lnCantidadSIS As Long, lnTotalSIS As Double
        Dim lnPrecioPagante  As Double, lnCantidadPagante As Long, lnTotalPagante As Double
        oConexion.Open sighEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        If sighEntidades.EsFecha(oDOAtencion.fechaEgreso, "DD/MM/AAAA") Then
            lcHoraEstanciaMax = lcBuscaParametro.SeleccionaFilaParametro(201)
            'Dias de Estancia
            lnDiasEstancia = oBuscaDiasPaciente.DiasDelPacienteEnHospitalizacionEmergencia(oDOAtencion.FechaIngreso, oDOAtencion.HoraIngreso, oDOAtencion.fechaEgreso, oDOAtencion.HoraEgreso, lcHoraEstanciaMax)
    
            'Horas de Estancia
            lnHorasEstancia = sighEntidades.HorasDelPacienteEnHospitalizacionEmergencia(oDOAtencion.FechaIngreso, oDOAtencion.HoraIngreso, oDOAtencion.fechaEgreso, oDOAtencion.HoraEgreso)
            
            'Codigo Servicio
            If oDOAtencion.idTipoServicio = sghHospitalizacion Then
               lnIdServicio = Val(wxParametro202)
            ElseIf lnHorasEstancia < 12 Then   'Emergencia Menor a 12 horas
               lnIdServicio = Val(wxParametro203)
            ElseIf lnHorasEstancia <= 24 Then  'Emergencia Entre 12 y 24 horas
               lnIdServicio = Val(wxParametro204)
            Else                               'Emergencia >24 horas
               lnIdServicio = Val(wxParametro204)
            End If
            oGenerarRecordsetProductos.GenerarRecordsetProductos oRsEstancia
            If lnIdServicio > 0 Then
                If oDOAtencion.idTipoServicio = sghHospitalizacion Then
                     mo_AdminAdmision.GeneraEstanciaPorCadaServicioHospitalizado oDOAtencion.idCuentaAtencion, _
                                      oDOAtencion.fechaEgreso, oDOAtencion.HoraEgreso, oRsEstancia, _
                                      lnTotalPagarEstancia, lnTotalDiasEstancia, oConexion, True, False
                     With oRsTmp
                        .Fields.Append "IdProducto", adInteger
                        .Fields.Append "CantidadEstancia", adInteger
                        .Fields.Append "PrecioEstancia", adCurrency
                        .Fields.Append "IdTipoFinanciamiento", adInteger
                        .Fields.Append "IdFuenteFinanciamiento", adInteger
                        .CursorType = adOpenDynamic
                        .LockType = adLockOptimistic
                        .Open
                     End With
                     If oRsEstancia.RecordCount > 0 Then
                        oRsEstancia.Sort = "idProducto"
                        oRsEstancia.MoveFirst
                        Do While Not oRsEstancia.EOF
                           lnIdProducto = oRsEstancia.Fields!idProducto
                           lnPrecioUnitario = oRsEstancia.Fields!PrecioEstancia
                           lnDiasEstancia = 0
                           lnPrecioSIS = oRsEstancia!precioSIS
                           lnCantidadSIS = 0
                           lnTotalSIS = 0
                           lnPrecioPagante = oRsEstancia!PrecioUnitario
                           lnCantidadPagante = 0
                           lnTotalPagante = 0
                           Do While Not oRsEstancia.EOF And lnIdProducto = oRsEstancia.Fields!idProducto
                                lnDiasEstancia = lnDiasEstancia + oRsEstancia.Fields!CantidadEstancia
                                If oDOAtencion.IdFormaPago = sghTipoFinanciamiento.sghSis Then
                                    lnCantidadSIS = lnCantidadSIS + oRsEstancia!CantidadSIS
                                    lnTotalSIS = lnTotalSIS + oRsEstancia!ImporteSIS
                                    lnCantidadPagante = lnCantidadPagante + oRsEstancia!Cantidad
                                    lnTotalPagante = lnTotalPagante + oRsEstancia!TotalPorPagar
                                End If
                                oRsEstancia.MoveNext
                                If oRsEstancia.EOF Then
                                   Exit Do
                                End If
                           Loop
                           If oDOAtencion.IdFormaPago = sghTipoFinanciamiento.sghSis Then
                                oRsTmp.AddNew
                                oRsTmp.Fields!idProducto = lnIdProducto
                                oRsTmp.Fields!CantidadEstancia = lnCantidadSIS
                                oRsTmp.Fields!PrecioEstancia = lnPrecioSIS
                                oRsTmp.Fields!idTipoFinanciamiento = sghTipoFinanciamiento.sghSis
                                oRsTmp.Fields!IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS
                                oRsTmp.Update
                                If lnCantidadPagante > 0 Then
                                     oRsTmp.AddNew
                                     oRsTmp.Fields!idProducto = lnIdProducto
                                     oRsTmp.Fields!CantidadEstancia = lnCantidadPagante
                                     oRsTmp.Fields!PrecioEstancia = lnPrecioPagante
                                     oRsTmp.Fields!idTipoFinanciamiento = sghTipoFinanciamiento.sghPacienteNormal
                                     oRsTmp.Fields!IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFParticularHospitalizado
                                     oRsTmp.Update
                                End If
                           Else
                                oRsTmp.AddNew
                                oRsTmp.Fields!idProducto = lnIdProducto
                                oRsTmp.Fields!CantidadEstancia = lnDiasEstancia
                                oRsTmp.Fields!PrecioEstancia = lnPrecioUnitario
                                oRsTmp.Fields!idTipoFinanciamiento = oDOAtencion.IdFormaPago
                                oRsTmp.Fields!IdFuenteFinanciamiento = oDOAtencion.IdFuenteFinanciamiento
                                oRsTmp.Update
                           End If
                        Loop
                        oRsEstancia.Close
                        Set oRsEstancia = oRsTmp
                     End If
                Else
                    '********cabecera/detalle
                     lnPrecioUnitario = 0
                     lnIdTipoFinanciamiento = oDOAtencion.IdFormaPago
                     oDoCatalogoServicioHosp.PrecioUnitario = 0
                     Set oDoCatalogoServicioHosp = mo_AdminFacturacion.CatalogoServiciosHospSeleccionarPorId(lnIdServicio, lnIdTipoFinanciamiento)
                     If oDoCatalogoServicioHosp.PrecioUnitario = 0 And oDoCatalogoServicioHosp.SeUsaSinPrecio = False Then
                        lcCodigoNombre = mo_AdminServiciosComunes.DevuelveNombreMedicamentoOServicioSegunId(lnIdServicio, sghServicio)
                        MsgBox "Tiene problemas con el ID SERVICIO: " & lcCodigoNombre & Chr(13) & "para el Producto/Plan: " & oDOAtencion.IdFormaPago & Chr(13) & "...consulte al ADMINISTRADOR DEL SISTEMA....", vbInformation, "Error"
                        Exit Function
                     End If
                     lnPrecioUnitario = oDoCatalogoServicioHosp.PrecioUnitario
                     oRsEstancia.AddNew
                     oRsEstancia.Fields!idProducto = lnIdServicio
                     oRsEstancia.Fields!CantidadEstancia = lnDiasEstancia
                     oRsEstancia.Fields!PrecioEstancia = lnPrecioUnitario
                     oRsEstancia.Fields!idTipoFinanciamiento = lnIdTipoFinanciamiento
                     oRsEstancia.Update
                End If
            End If
        End If
        CalculaEstanciaParaPacienteConAltaMedica = True
        Set oRsTmp = Nothing
        Set oDoCatalogoServicioHosp = Nothing
        Set oGenerarRecordsetProductos = Nothing
        oConexion.Close
        Set oConexion = Nothing
End Function

Sub InicilizarParametros()
    wxParametro202 = lcBuscaParametro.SeleccionaFilaParametro(202)
    wxParametro203 = lcBuscaParametro.SeleccionaFilaParametro(203)
    wxParametro204 = lcBuscaParametro.SeleccionaFilaParametro(204)
    wxParametro208 = lcBuscaParametro.SeleccionaFilaParametro(208)
    wxParametro210 = lcBuscaParametro.SeleccionaFilaParametro(210)
    wxParametro212 = lcBuscaParametro.SeleccionaFilaParametro(212)
    wxParametro215 = lcBuscaParametro.SeleccionaFilaParametro(215)
    wxParametro216 = lcBuscaParametro.SeleccionaFilaParametro(216)
    wxParametro231 = lcBuscaParametro.SeleccionaFilaParametro(231)
    wxParametro232 = lcBuscaParametro.SeleccionaFilaParametro(232)
    wxParametro233 = lcBuscaParametro.SeleccionaFilaParametro(233)
    wxParametro237 = lcBuscaParametro.SeleccionaFilaParametro(237)
    wxParametro259 = lcBuscaParametro.SeleccionaFilaParametro(259)
    wxParametro290 = lcBuscaParametro.SeleccionaFilaParametro(290)
    wxParametro291 = lcBuscaParametro.SeleccionaFilaParametro(291)
    wxParametro292 = lcBuscaParametro.SeleccionaFilaParametro(292)
    wxParametro296 = lcBuscaParametro.SeleccionaFilaParametro(296)
    wxParametro302 = lcBuscaParametro.SeleccionaFilaParametro(302)
    wxParametro306 = lcBuscaParametro.SeleccionaFilaParametro(306)
    wxParametro312 = lcBuscaParametro.SeleccionaFilaParametro(312)
    wxParametro316 = lcBuscaParametro.SeleccionaFilaParametro(316): wxParametro316 = IIf(IsNull(wxParametro316), "", wxParametro316)
    wxParametro317 = Trim(lcBuscaParametro.SeleccionaFilaParametro(317)): wxParametro317 = IIf(IsNull(wxParametro317), "", wxParametro317)
    wxParametro322 = lcBuscaParametro.SeleccionaFilaParametro(322)
    wxParametro323 = lcBuscaParametro.SeleccionaFilaParametro(323)
    wxParametro324 = lcBuscaParametro.SeleccionaFilaParametro(324)
    wxParametro358 = lcBuscaParametro.SeleccionaFilaParametro(358)
    wxParametro359 = lcBuscaParametro.SeleccionaFilaParametro(359)
    wxParametro511 = lcBuscaParametro.SeleccionaFilaParametro(511)
    wxParametro512 = lcBuscaParametro.SeleccionaFilaParametro(512)
    wxParametro530 = lcBuscaParametro.SeleccionaFilaParametro(530)
    wxParametro552 = lcBuscaParametro.SeleccionaFilaParametro(552)
    wxParametro559 = lcBuscaParametro.SeleccionaFilaParametro(559)
    wxParametroSIS = lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
    ldFechaActualServidor = lcBuscaParametro.RetornaFechaServidorSQL
    wxParametroJAMO = lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
End Sub




Sub HaceVisibleOnoBotonFUA()
    If wxParametro302 = "S" Then
        btnImprimeFichaSIS.Visible = False
        If ml_IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
           btnImprimeFichaSIS.Visible = True
           Dim lcSexo As String, ml_edad_En_YYYYMMDD As String
           ml_edad_En_YYYYMMDD = sighEntidades.EdadActualEnFormatoYYYYMMDD(mo_Pacientes.FechaNacimiento, _
                                    CDate(Format(ml_ldFechaIngreso, sighEntidades.DevuelveFechaSoloFormato_DMY) & " " & ml_lcHoraIngreso))
           lcSexo = IIf(mo_Pacientes.idTipoSexo = 1, "M", "F")
        End If
    End If
End Sub

'mgaray09
Private Function setListItemAControlDiagnosticos(IdListBarItem As Long)
    ucDiagnosticosEgreso.IdListBarItem = IdListBarItem
    ucDiagnosticoComplicaciones.IdListBarItem = IdListBarItem
    ucDiagnosticoNacimiento.IdListBarItem = IdListBarItem
    ucDiagnosticosMortalidad.IdListBarItem = IdListBarItem
End Function
'mgaray201410e
Private Function ServicioImprimeFUAAdmision() As Boolean
On Error GoTo miError
    Dim oDoServicio As New doServicio
    Dim returnValue As Boolean
    Dim oDOCama As New DOCama
    Dim oConexion As New Connection
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    returnValue = False
    
    Set oDoServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(mo_Atenciones.IdServicioIngreso, oConexion)
    If Not oDoServicio Is Nothing Then
        If oDoServicio.UsaGalenHos = True Then
            returnValue = True
        End If
    End If
    
    ServicioImprimeFUAAdmision = returnValue
    oConexion.Close
    Set oConexion = Nothing
miError:
    If Err Then
        MsgBox Err.Number & ". " & Err.Description, vbCritical, "Egreso"
    End If
    
End Function


Sub CreaCuentaAutomaticaEnHospitalizacion()
           'debb-29/05/2018
           Dim oRsTmp1 As New Recordset
           Set oRsTmp1 = mo_AdminAdmision.AtencionesDatosAdicionalesSeleccionarPorIdPaciente(mo_Atenciones.idPaciente)
           oRsTmp1.Filter = "idAtencionEmeg_CE=" & mo_Atenciones.idAtencion
           If oRsTmp1.RecordCount > 0 Then
            oRsTmp1.Close
            Set oRsTmp1 = Nothing
            Exit Sub
           End If
           oRsTmp1.Close
           Set oRsTmp1 = Nothing
           '
           Dim ml_idAtencionEmeg_CE As Long
           'lblNombreMedico.Text = ""
           ml_TipoServicio = sghHospitalizacion
           'mo_cmbIdTipoServicio.BoundText = ml_TipoServicio
           
           ml_idAtencionEmeg_CE = mo_Atenciones.idAtencion
           'Me.txtHoraIngreso.Text = Me.txtHoraEgreso.Text
           'Me.txtFechaIngreso.Text = Me.txtFechaEgreso.Text
           'txtHoraEgreso = sighentidades.HORA_VACIA_HM
           mo_cmbCondicionAlta.BoundText = 0
           mi_Opcion = sghAgregar
           CargaDatosAlObjetosDeDatos True
           mo_Atenciones.FechaIngreso = Me.txtFechaEgreso.Text
           mo_Atenciones.FechaEgresoAdministrativo = 0
           mo_Atenciones.fechaEgreso = 0
           mo_Atenciones.HoraIngreso = Me.txtHoraEgreso.Text
           mo_Atenciones.HoraEgreso = ""
           mo_Atenciones.HoraEgresoAdministrativo = ""
           mo_Atenciones.idAtencion = 0
           mo_Atenciones.IdCamaEgreso = 0
           mo_Atenciones.IdCamaIngreso = 0
           mo_Atenciones.IdCondicionAlta = 0
           mo_Atenciones.idCuentaAtencion = 0
           mo_Atenciones.IdDestinoAtencion = 0
           mo_Atenciones.IdEspecialidadMedico = 0
           mo_Atenciones.IdEstadoAtencion = 1
           mo_Atenciones.IdMedicoEgreso = 0
           mo_Atenciones.IdMedicoIngreso = 0
           mo_Atenciones.IdServicioIngreso = Val(mo_cmbServicioDestino.BoundText) '282     'adicciones
           mo_Atenciones.IdServicioEgreso = mo_Atenciones.IdServicioIngreso
           mo_Atenciones.IdTipoAlta = 0
           'mo_Atenciones.IdTipoCondicionALEstab = 0
           'mo_Atenciones.IdTipoCondicionAlServicio = 0
           mo_Atenciones.IdTipoGravedad = 0
           mo_Atenciones.idTipoServicio = sghHospitalizacion
           mo_Atenciones.PisoDomicilio = ""
           mo_Atenciones.IdOrigenAtencion = 31
           mo_CuentasAtencion.idEstado = sghEstadoCuenta.sghAbierto
           mo_CuentasAtencion.FechaApertura = mo_Atenciones.FechaIngreso
           mo_CuentasAtencion.HoraApertura = mo_Atenciones.HoraIngreso
           mo_DoAtencionDatosAdicionales.idAtencionEmeg_CE = ml_idAtencionEmeg_CE
           
           If mo_AdminAdmision.AdmisionHospAgregar(mo_CuentasAtencion, mo_Atenciones, mo_Pacientes, mo_Historia, _
                                                        0, mo_OcupacionCamas, _
                                                        mo_Diagnosticos, mo_Procedimientos, mo_Examenes, mo_Nacimientos, _
                                                        0, _
                                                        mo_AtencionesEmergencia, mo_AtencionPadre, lbPacienteNN, _
                                                        mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, _
                                                        "Creación de Cuenta automática desde Emergencia", _
                                                        lnIdNacimientoSeleccionado, oDoSunasaPacientesHistoricos, _
                                                        mo_DoAtencionDatosAdicionales, False) = True Then
                ImprimePreCuenta
           End If
End Sub

Sub ImprimePreCuenta()
    Dim oReporte As New RptCaja
    Dim lcPaciente As String
    Dim lcMedico As String
    lcPaciente = Trim(mo_Pacientes.ApellidoPaterno) & " " & Trim(mo_Pacientes.ApellidoMaterno) & " " & Trim(mo_Pacientes.PrimerNombre)
    If mo_Pacientes.SegundoNombre <> "" Then
       lcPaciente = lcPaciente & " " & Trim(mo_Pacientes.SegundoNombre)
    End If
    If mo_Pacientes.TercerNombre <> "" Then
      lcPaciente = lcPaciente & " " & Trim(mo_Pacientes.TercerNombre)
    End If
    lcMedico = "..." 'lblNombreMedicoEgreso.Text

    oReporte.ImpresionPreCuenta Format(mo_Atenciones.FechaIngreso, sighEntidades.DevuelveFechaSoloFormato_DMY), mo_Atenciones.FechaIngreso, lcPaciente, mo_Pacientes.NroHistoriaClinica, _
                                cmbServicioDestino.Text, lcMedico, IIf(ml_TipoServicio = sghHospitalizacion, _
                                "HOSPITALIZACION", "EMERGENCIA"), mo_Atenciones.idAtencion, "", _
                                mo_Atenciones.idCuentaAtencion, ml_TipoFinanciamiento, "", ml_idUsuario, _
                                "Cama: ", mo_Pacientes.FichaFamiliar, mo_Pacientes.idTipoNumeracion, _
                                wxParametro216, wxParametro306, False, 0
    Set oReporte = Nothing
'    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub



Private Sub cmbTipoAtencion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbTipoAtencion
    AdministrarKeyPreview KeyCode
End Sub
Private Sub cmbComoLlego_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbComoLlego
    AdministrarKeyPreview KeyCode

End Sub
Private Sub cmbEstadoLlegada_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbEstadoLlegada
    AdministrarKeyPreview KeyCode
End Sub
Private Sub cmbIdTipoGravedad_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoGravedad
    AdministrarKeyPreview KeyCode
End Sub


