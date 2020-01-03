VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FacGeneraCtaPacienteExtSeguro 
   Caption         =   "Genera Cuenta para un Paciente Externo - CON SEGURO"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10710
   Icon            =   "FacGeneraCtaPacienteExtSeguro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   10710
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   30
      TabIndex        =   8
      Top             =   6420
      Width           =   10605
      Begin SIGHProxies.ucMensajeParpadeando ucMensajeParpadeando1 
         Height          =   675
         Left            =   1620
         TabIndex        =   58
         Top             =   270
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   1191
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
         Height          =   700
         Left            =   9090
         Picture         =   "FacGeneraCtaPacienteExtSeguro.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   240
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.CommandButton btnImprimePreCta 
         Caption         =   "Imprime Cuenta"
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
         Picture         =   "FacGeneraCtaPacienteExtSeguro.frx":04E5
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   225
         Width           =   1245
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "FacGeneraCtaPacienteExtSeguro.frx":09BE
         DownPicture     =   "FacGeneraCtaPacienteExtSeguro.frx":0E1E
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
         Left            =   3900
         Picture         =   "FacGeneraCtaPacienteExtSeguro.frx":1293
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "FacGeneraCtaPacienteExtSeguro.frx":1708
         DownPicture     =   "FacGeneraCtaPacienteExtSeguro.frx":1BCC
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
         Left            =   5438
         Picture         =   "FacGeneraCtaPacienteExtSeguro.frx":20B8
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   225
         Width           =   1365
      End
   End
   Begin UltraGrid.SSUltraGrid grdPacientesEncontrados 
      Height          =   225
      Left            =   30
      TabIndex        =   7
      Top             =   810
      Visible         =   0   'False
      Width           =   5505
      _ExtentX        =   9710
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
      Caption         =   "Lista de pacientes encontrados"
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5355
      Left            =   30
      TabIndex        =   10
      Top             =   1020
      Width           =   10665
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
         Height          =   1830
         Left            =   4605
         TabIndex        =   42
         Top             =   3480
         Width           =   6000
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
            Left            =   1890
            TabIndex        =   50
            Top             =   300
            Width           =   1455
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
            TabIndex        =   49
            Top             =   660
            Width           =   2820
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
            Left            =   1905
            MaxLength       =   8
            TabIndex        =   48
            Top             =   660
            Width           =   750
         End
         Begin VB.CommandButton btnBuscarEstablecimiento 
            Caption         =   "..."
            Height          =   315
            Left            =   2700
            TabIndex        =   47
            Top             =   660
            Width           =   315
         End
         Begin VB.TextBox txtReferenciaO 
            Height          =   315
            Left            =   4635
            MaxLength       =   20
            TabIndex        =   46
            Top             =   285
            Width           =   1230
         End
         Begin VB.CommandButton btnDxReferencia 
            Caption         =   "..."
            Height          =   315
            Left            =   2700
            TabIndex        =   45
            Top             =   1350
            Width           =   315
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
            Left            =   1905
            MaxLength       =   8
            TabIndex        =   44
            Top             =   1350
            Width           =   750
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
            TabIndex        =   43
            Top             =   1350
            Width           =   2820
         End
         Begin PVCOMBOLibCtl.PVComboBox cmbServicioReferenciaO 
            Height          =   330
            Left            =   1905
            TabIndex        =   51
            Top             =   1005
            Width           =   3960
            _Version        =   524288
            _cx             =   6985
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
            TabIndex        =   56
            Top             =   360
            Width           =   1230
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
            TabIndex        =   55
            Top             =   705
            Width           =   1380
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
            TabIndex        =   54
            Top             =   315
            Width           =   1125
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
            TabIndex        =   53
            Top             =   1035
            Width           =   1485
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
            TabIndex        =   52
            Top             =   1395
            Width           =   1080
         End
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   5130
         Left            =   45
         TabIndex        =   37
         Top             =   180
         Width           =   4530
         Begin SIGHProxies.UcPacienteDatosAloj UcPacienteDatosAloj1 
            Height          =   3660
            Left            =   45
            TabIndex        =   57
            Top             =   195
            Width           =   4440
            _ExtentX        =   7832
            _ExtentY        =   6456
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Atención"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   4620
         TabIndex        =   17
         Top             =   300
         Width           =   5985
         Begin VB.Frame Frame5 
            Height          =   30
            Left            =   0
            TabIndex        =   29
            Top             =   1830
            Width           =   5985
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
            Left            =   1890
            TabIndex        =   25
            Top             =   2670
            Width           =   2235
         End
         Begin VB.TextBox txtIdMedicoIngreso 
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
            Left            =   4920
            TabIndex        =   20
            Top             =   600
            Width           =   915
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
            ItemData        =   "FacGeneraCtaPacienteExtSeguro.frx":25A4
            Left            =   2550
            List            =   "FacGeneraCtaPacienteExtSeguro.frx":25A6
            TabIndex        =   19
            Top             =   1380
            Width           =   1545
         End
         Begin VB.ComboBox cmbServicioIngreso 
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
            Left            =   1890
            TabIndex        =   11
            Top             =   210
            Width           =   3960
         End
         Begin VB.TextBox lblNombreMedico 
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
            Left            =   1890
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   600
            Width           =   3015
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
            Left            =   1890
            TabIndex        =   18
            Top             =   1380
            Width           =   585
         End
         Begin MSMask.MaskEdBox txtHoraIngreso 
            Height          =   315
            Left            =   3300
            TabIndex        =   14
            Top             =   990
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
         Begin MSMask.MaskEdBox txtFechaIngreso 
            Height          =   315
            Left            =   1890
            TabIndex        =   13
            Top             =   990
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
         Begin MSDataListLib.DataCombo cmbFuenteFinanciamiento 
            Height          =   330
            Left            =   1890
            TabIndex        =   15
            Top             =   1890
            Width           =   3945
            _ExtentX        =   6959
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo cmbFormaPago 
            Height          =   330
            Left            =   1890
            TabIndex        =   35
            Top             =   2280
            Width           =   3945
            _ExtentX        =   6959
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   "DataCombo1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Producto/Plan"
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
            TabIndex        =   36
            Top             =   2370
            Width           =   1155
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Fte.Financiam/IAFA"
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
            TabIndex        =   28
            Top             =   1950
            Width           =   1575
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
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
            Height          =   210
            Left            =   120
            TabIndex        =   27
            Top             =   2730
            Width           =   855
         End
         Begin VB.Label lblEstadoCta 
            AutoSize        =   -1  'True
            Caption         =   "."
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
            Left            =   4140
            TabIndex        =   26
            Top             =   2730
            Width           =   1680
         End
         Begin VB.Label lblIdMedicoIngreso 
            Caption         =   "Responsable"
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
            Top             =   645
            Width           =   1335
         End
         Begin VB.Label lblIdServicioIngreso 
            Caption         =   "Servicio ingreso"
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
            TabIndex        =   23
            Top             =   300
            Width           =   1395
         End
         Begin VB.Label lblEdadEnDias 
            Caption         =   "Edad en Atención"
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
            TabIndex        =   22
            Top             =   1410
            Width           =   1725
         End
         Begin VB.Label lblFecha 
            Caption         =   "Fecha ingreso"
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
            TabIndex        =   21
            Top             =   1020
            Width           =   1215
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
      Height          =   975
      Left            =   60
      TabIndex        =   30
      Top             =   0
      Width           =   10695
      Begin VB.CommandButton cmdSinApellidoMaterno 
         Caption         =   "..."
         Height          =   315
         Left            =   2880
         TabIndex        =   41
         ToolTipText     =   "Sin apellido MATERNO"
         Top             =   450
         Width           =   255
      End
      Begin VB.CommandButton cmdSinApellidoPaterno 
         Caption         =   "..."
         Height          =   315
         Left            =   1830
         TabIndex        =   40
         ToolTipText     =   "Sin apellido PATERNO"
         Top             =   450
         Width           =   255
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   8040
         Picture         =   "FacGeneraCtaPacienteExtSeguro.frx":25A8
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   210
         Width           =   1305
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
         Left            =   3210
         MaxLength       =   40
         TabIndex        =   2
         Top             =   450
         Width           =   600
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
         Left            =   930
         MaxLength       =   40
         TabIndex        =   0
         Top             =   450
         Width           =   885
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
         Left            =   2130
         MaxLength       =   40
         TabIndex        =   1
         Top             =   450
         Width           =   750
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
         Left            =   3840
         MaxLength       =   40
         TabIndex        =   3
         Top             =   450
         Width           =   570
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
         Left            =   4440
         TabIndex        =   4
         Top             =   450
         Width           =   885
      End
      Begin VB.CommandButton btnBuscarPaciente 
         Height          =   315
         Left            =   8040
         Picture         =   "FacGeneraCtaPacienteExtSeguro.frx":2BD1
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   540
         Width           =   1305
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
         Left            =   120
         TabIndex        =   6
         Top             =   450
         Width           =   795
      End
      Begin VB.Frame fraPacienteNuevo 
         Height          =   795
         Left            =   9360
         TabIndex        =   31
         Top             =   120
         Visible         =   0   'False
         Width           =   1245
         Begin VB.CheckBox chkPacienteNuevo 
            Alignment       =   1  'Right Justify
            Caption         =   "Paciente &nuevo"
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
            Left            =   30
            TabIndex        =   32
            Top             =   180
            Width           =   1095
         End
      End
      Begin VB.Label Label50 
         Caption         =   "Nº Historia  Apelli.Paterno   ApelliMaterno  1erNom  2oNom  DNI"
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
         Left            =   180
         TabIndex        =   33
         Top             =   240
         Width           =   4575
      End
   End
End
Attribute VB_Name = "FacGeneraCtaPacienteExtSeguro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento para generar CUENTA a un Paciente con Seguro
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mi_Opcion As sghOpciones
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Apariencia As New sighentidades.GridInfragistic
'
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminProgramacion As New SIGHNegocios.ReglasDeProgMedica
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasSISgalenhos As New ReglasSISgalenhos
'
Dim mo_cmbServicioIngreso As New sighentidades.ListaDespleglable
Dim mo_cmbIdTipoEdad As New sighentidades.ListaDespleglable
Dim mo_cmbIdTipoReferenciaOrigen As New sighentidades.ListaDespleglable
'
Dim oRsFuentesFinanciamiento As New Recordset
Dim oRsFormaPago As New Recordset
'
Dim mo_CuentasAtencion As New DOCuentaAtencion
Dim mo_Atenciones As New DOAtencion
Dim mo_Pacientes  As New doPaciente
Dim mo_DoAtencionDatosAdicionales As New DoAtencionDatosAdicionales
'
Dim ml_idCuentaAtencion As Long
Dim ml_idPaciente As Long
'
Dim lcBuscaParametro As New SIGHDatos.Parametros
'
Dim lcSql As String
Dim ms_MensajeError As String
Dim lnEspecialidadServicio As Long
Dim lbUltimaTeclaPulsoENTER As Boolean
'
Dim ml_lbPacienteTieneSeguro As Boolean
Dim ml_idUsuario As Long
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim ml_idAtencion As Long
Dim lnAfiliacionSIS4 As Long, lcSIScodigo As String
Dim ldFechaActualServidor As Date
Dim lcDNI_busqueda As String
Dim lcCodigoEstablecimientoAdscripcionSIS As String

Property Let lcNombrePc(lValue As String)
  mo_lcNombrePc = lValue
End Property

Property Let lnIdTablaLISTBARITEMS(lValue As Long)
  mo_lnIdTablaLISTBARITEMS = lValue
End Property

Property Let idUsuario(lValue As Long)
  ml_idUsuario = lValue
End Property

Property Let Opcion(iValue As sghOpciones)
  mi_Opcion = iValue
End Property

Property Get Opcion() As sghOpciones
  Opcion = mi_Opcion
End Property

Property Let lbPacienteTieneSeguro(lValue As Boolean)
  ml_lbPacienteTieneSeguro = lValue
End Property

Property Let idAtencion(lValue As Long)
  ml_idAtencion = lValue
End Property

Property Get idAtencion() As Long
  idAtencion = ml_idAtencion
End Property


Sub LimpiarBusqueda()
    Me.txtNroHistoriaBusqueda.Text = ""
    Me.txtApellidoPaternoBusqueda.Text = ""
    Me.txtApellidoMaternoBusqueda.Text = ""
    Me.txtPrimerNombreBusqueda.Text = ""
    Me.txtSegundoNombreBusqueda.Text = ""
    Me.txtNroDNIBusqueda.Text = ""

End Sub

Private Sub btnAceptar_Click()
  If btnAceptar.Enabled = False Then Exit Sub
  Select Case mi_Opcion
    Case sghAgregar
      If ValidarDatosObligatorios() Then
        CargaDatosAlObjetosDeDatos
        If ValidarReglas() Then
          If AgregarDatos() Then
                Me.txtNroCuenta = mo_Atenciones.idCuentaAtencion
                MsgBox " Los datos se agregaron correctamente, para la Historia Nª: " & mo_Pacientes.NroHistoriaClinica & Chr(13) & Chr(13) & "N° Cuenta " & txtNroCuenta.Text, vbInformation, Me.Caption
                If Me.txtNroCuenta.Text <> "" Then
                   ImprimePreCuenta
                   Me.Visible = False
                End If
                If wxParametro302 = "S" And mo_Atenciones.IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                     btnImprimeFichaSIS_Click
                End If
                Me.Visible = False
          Else
            MsgBox "No se pudo agregar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
          End If
        End If
      End If
    Case sghModificar
       If ValidarDatosObligatorios() Then
         CargaDatosAlObjetosDeDatos
         If ValidarReglas() Then
           If ModificarDatos() Then
                MsgBox " Los datos se modificaron correctamente, para la Cuenta N° " & txtNroCuenta.Text, vbInformation, Me.Caption
                If wxParametro302 = "S" And mo_Atenciones.IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                     btnImprimeFichaSIS_Click
                End If
                Me.Visible = False
           Else
                MsgBox "No se pudo modificar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
           End If
         End If
       End If
    Case sghEliminar
      If ValidarReglas() Then
        CargaDatosAlObjetosDeDatos
        If EliminarDatos() Then
          MsgBox "Los datos se eliminaron correctamente, para la Cuenta N° " & txtNroCuenta.Text, vbInformation, Me.Caption
          Me.Visible = False
        Else
          MsgBox "No se pudo eliminar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
        End If
      End If
  End Select
End Sub


Function ValidarDatosObligatorios() As Boolean
   Dim sMensaje As String
   ValidarDatosObligatorios = False
   UcPacienteDatosAloj1.CargarDatosAlObjetoDatos mo_Pacientes
   If mo_Pacientes.ApellidoPaterno = "" Then
       sMensaje = sMensaje + "Ingrese el Apellido Paterno " + Chr(13)
   End If
   If mo_Pacientes.ApellidoMaterno = "" Then
       sMensaje = sMensaje + "Ingrese el Apellido Materno " + Chr(13)
   End If
   If mo_Pacientes.PrimerNombre = "" Then
       sMensaje = sMensaje + "Ingrese el Apellido Primer Nombre" + Chr(13)
   End If
   If mo_Pacientes.idTipoSexo = 0 Then
       sMensaje = sMensaje + "Elija el Sexo" + Chr(13)
   End If
   If cmbServicioIngreso.Text = "" Then
       sMensaje = sMensaje + "Elija el Servicio de Ingreso" + Chr(13)
   End If
   If txtIdMedicoIngreso.Text = "" Then
       sMensaje = sMensaje + "Ingrese el Responsable" + Chr(13)
   End If
   If txtFechaIngreso.Text = sighentidades.FECHA_VACIA_DMY Then
       sMensaje = sMensaje + "Registre la Fecha de Ingreso " + Chr(13)
   End If
   If txtHoraIngreso.Text = sighentidades.HORA_VACIA_HM Then
       sMensaje = sMensaje + "Registre la Hora de Ingreso" + Chr(13)
   End If
   If txtEdadEnDias.Text = "" Then
       sMensaje = sMensaje + "Ingrese la Edad" + Chr(13)
   End If
   If cmbIdTipoEdad.Text = "" Then
       sMensaje = sMensaje + "Elija el Tipo de Edad" + Chr(13)
   End If
   If Me.cmbFuenteFinanciamiento.Text = "" Then
      sMensaje = sMensaje + "Elija el Plan de Atención" + Chr(13)
   End If
   If Not Val(cmbFormaPago.BoundText) > 0 Then
      sMensaje = sMensaje + "Elija el Tipo de Financiamiento" + Chr(13)
   End If
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
    ValidarReglas = False
    Dim rsCitas As New Recordset
    Dim lcMensaje As String
    
    If mi_Opcion = sghAgregar And mo_Pacientes.idPaciente > 0 Then
             lcMensaje = mo_AdminFacturacion.DevuelveSiElPacienteFallecioOhistoriaPasoPasivo(mo_Pacientes.idPaciente)
             If lcMensaje <> "" Then
                MsgBox lcMensaje, vbInformation, Me.Caption
                Exit Function
             End If
    End If
    
    '
    ValidarReglas = UcPacienteDatosAloj1.ValidarReglas(mo_Pacientes)
    '
    If wxParametro302 = "S" And Val(cmbFuenteFinanciamiento.BoundText) = sghFuenteFinanciamiento.sghFFSIS And _
            mi_Opcion = sghEliminar Then
            Set rsCitas = mo_ReglasSISgalenhos.SisFuaAtencionSeleccionarPorCuenta(Val(Me.txtNroCuenta.Text))
            If rsCitas.RecordCount > 0 Then
               MsgBox "El formato FUA ya fué generado: " & rsCitas.Fields!fuaDisa & "-" & rsCitas!fuaLote & "-" & _
                      rsCitas!FuaNumero & Chr(13) & "Debe eliminar el formato FUA (módulo: SIS, opción: Formato FUA)", _
                      vbInformation, Me.Caption
               Set rsCitas = Nothing
               ValidarReglas = False
               Exit Function
            End If
    End If
    'DEBB-04/11/2016
    If txtNombreOrigenReferencia.Text <> "" Then
       If Val(txtReferenciaO.Text) = 0 Then
          MsgBox "Debe registrar el N° REFERENCIA", vbInformation, Me.Caption
          ValidarReglas = False
          Exit Function
       ElseIf cmbServicioReferenciaO.Text = "" Then
          MsgBox "Debe registrar SERVICIO REFERENCIA", vbInformation, Me.Caption
          ValidarReglas = False
          Exit Function
       ElseIf lblDxReferencia1.Text = "" Then
          MsgBox "Debe registrar DX REFERENCIA", vbInformation, Me.Caption
          ValidarReglas = False
          Exit Function
       End If
    End If
    
    Set rsCitas = Nothing
    ValidarReglas = True
End Function

Sub CargaDatosAlObjetosDeDatos()
    Dim oRsTmp As New Recordset
    Dim lnIdTipoServicio As Long
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DEL PACIENTE
    '---------------------------------------------------------------------------------
    '********mo_Pacientes****** YA SE CARGO EN VALIDADATOSOBLIGATORIOS()
    '
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LA CUENTA ATENCION
    '---------------------------------------------------------------------------------
    Select Case mi_Opcion
    Case sghAgregar
        With mo_CuentasAtencion
                .idPaciente = mo_Pacientes.idPaciente
                .TotalAsegurado = 0
                .TotalExonerado = 0
                .TotalPagado = 0
                .TotalPorPagar = 0
                .idEstado = sghEstadoCuenta.sghAbierto
                .FechaApertura = Me.txtFechaIngreso.Text
                .HoraApertura = Me.txtHoraIngreso.Text
                .fechaCierre = 0
                .HoraCierre = ""
                .IdUsuarioAuditoria = ml_idUsuario
        End With
    Case Else
    End Select
   
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LA ATENCION
    '---------------------------------------------------------------------------------
    lnIdTipoServicio = 0
    Set oRsTmp = mo_AdminServiciosComunes.ServiciosSeleccionarXidentificador(mo_cmbServicioIngreso.BoundText)
    If oRsTmp.RecordCount > 0 Then
       lnIdTipoServicio = oRsTmp.Fields!idTipoServicio
    End If
    oRsTmp.Close
    With mo_Atenciones
           .idAtencion = Me.idAtencion
           .IdEspecialidadMedico = 0
           .IdMedicoIngreso = Val(Me.txtIdMedicoIngreso.Tag)
           .IdServicioIngreso = Val(mo_cmbServicioIngreso.BoundText)
           .HoraIngreso = IIf(Me.txtHoraIngreso.Text = sighentidades.HORA_VACIA_HM, "", Me.txtHoraIngreso.Text)
           .FechaIngreso = IIf(Me.txtFechaIngreso.Text = sighentidades.HORA_VACIA_HM, "", Me.txtFechaIngreso.Text)
           .idTipoServicio = lnIdTipoServicio
           .Edad = Me.txtEdadEnDias.Text
           .IdTipoEdad = Val(mo_cmbIdTipoEdad.BoundText)
           .idPaciente = mo_Pacientes.idPaciente
           .IdUsuarioAuditoria = ml_idUsuario
            .IdTipoCondicionALEstab = 1
            .IdTipoCondicionAlServicio = 1
            .FechaEgreso = 0
            .HoraEgreso = sighentidades.HORA_VACIA_HM
            .IdCondicionAlta = 0
            .IdTipoAlta = 0
            .IdTipoGravedad = 0
            .IdFormaPago = Val(Me.cmbFormaPago.BoundText)
            .IdFuenteFinanciamiento = Val(cmbFuenteFinanciamiento.BoundText)
            .idCuentaAtencion = Val(Me.txtNroCuenta.Text)
            .IdEstadoAtencion = 1
            .EsPacienteExterno = True
            
   End With
   Set oRsTmp = Nothing
   '
   With mo_DoAtencionDatosAdicionales
       .DireccionDomicilio = mo_Pacientes.DireccionDomicilio
       '.FuaCodigoPrestacion
       '.HuboInfeccionIntraHospitalaria
       .idAtencion = mo_Atenciones.idAtencion
       '.IdEstablecimientoDestino
       '.IdEstablecimientoNoMinsaDestino
       '.IdEstablecimientoNoMinsaOrigen
       '.IdEstablecimientoOrigen
       '.IdMedicoRespNacimiento
       .idSiasis = lnAfiliacionSIS4
       '.IdTipoReferenciaDestino
       '.IdTipoReferenciaOrigen
       .IdUsuarioAuditoria = mo_Pacientes.IdUsuarioAuditoria
       '.NombreAcompaniante
       '.NroReferenciaDestino
       '.NroReferenciaOrigen
       '.NumeroDeHijos
       '.Observacion
       '.ProximaCita
       '.RecienNacido
       .SisCodigo = lcSIScodigo
       '.TieneNecropsia
       
       'debb-04/11/2016
        .IdTipoReferenciaOrigen = Val(mo_cmbIdTipoReferenciaOrigen.BoundText)
        If .IdTipoReferenciaOrigen = 1 Then
            .IdEstablecimientoOrigen = Val(Me.txtIdEstablecimientoOrigen.Tag)
            .IdEstablecimientoNoMinsaOrigen = 0
        Else
            .IdEstablecimientoOrigen = 0
            .IdEstablecimientoNoMinsaOrigen = Val(Me.txtIdEstablecimientoOrigen.Tag)
        End If
        .NroReferenciaOrigen = txtReferenciaO.Text
        .referenciaOservicio = PVcomboBoxDevuelveEleccion(cmbServicioReferenciaO)
        .referenciaOidDiagnostico = Val(txtDxReferencia.Tag)
       
   End With
   '

End Sub


Private Sub btnBuscarEstablecimiento_Click()
    If cmbIdTipoReferenciaOrigen.Text <> "" Then
       CompletarDatosDeEstablecimiento txtIdEstablecimientoOrigen, txtNombreOrigenReferencia, mo_cmbIdTipoReferenciaOrigen.BoundText
    End If

End Sub

Private Sub btnBuscarPaciente_Click()
    Dim RsHistorias As New Recordset
    Dim oDOPaciente As New doPaciente
    
    On Error GoTo ErrBusq
    If mo_Teclado.TextoEsSoloNumeros(Me.txtNroHistoriaBusqueda.Text) Then
       oDOPaciente.NroHistoriaClinica = Val(Me.txtNroHistoriaBusqueda.Text)
    End If
    oDOPaciente.ApellidoPaterno = Me.txtApellidoPaternoBusqueda
    oDOPaciente.ApellidoMaterno = Me.txtApellidoMaternoBusqueda
    oDOPaciente.PrimerNombre = Me.txtPrimerNombreBusqueda
    oDOPaciente.SegundoNombre = Me.txtSegundoNombreBusqueda
    oDOPaciente.IdDocIdentidad = 1
    oDOPaciente.NroDocumento = Me.txtNroDNIBusqueda
    '
    lcDNI_busqueda = Trim(txtNroDNIBusqueda.Text)
    '
    If (oDOPaciente.ApellidoPaterno + oDOPaciente.ApellidoMaterno + _
    oDOPaciente.PrimerNombre + oDOPaciente.SegundoNombre = "") And _
    (Val(Me.txtNroHistoriaBusqueda.Text) = 0) And _
    (oDOPaciente.NroDocumento = "") Then
        MsgBox "Ingrese alguno de los valores de búsqueda", vbInformation, Me.Caption
        Exit Sub
    End If
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    Set RsHistorias = mo_AdminAdmision.PacientesFiltrarTodosSoloHistoriasDefinitivas(oDOPaciente, wxSinApellido, oConexion)
    Set grdPacientesEncontrados.DataSource = RsHistorias
    With grdPacientesEncontrados
        .Left = 240
        .Top = 780
        .Width = 11775
        .Height = 4455
    End With
    ml_idPaciente = 0
    'Si hay una sola coincidencia
    If RsHistorias.RecordCount = 1 Then
        If mo_AdminAdmision.BuscaSiEstaHospitalizado(RsHistorias!idPaciente, oConexion, sghConsultaExterna) = False Then  'debb-05/12/2015
            Me.grdPacientesEncontrados.Visible = False
            RsHistorias.MoveFirst
            chkPacienteNuevo.Value = 0
            ml_idPaciente = RsHistorias!idPaciente
            UcPacienteDatosAloj1.idPaciente = RsHistorias!idPaciente
            UcPacienteDatosAloj1.CargarDatosDePacienteALosControles
            CalculaEdadEnLaAtencion
            DeudasPendientesDeAnterioresAtenciones oConexion
            LimpiarBusqueda
            cmbServicioIngreso.SetFocus
        End If
    ElseIf RsHistorias.RecordCount > 1 Then
        Me.grdPacientesEncontrados.Visible = True
    ElseIf RsHistorias.RecordCount = 0 Then
        Me.grdPacientesEncontrados.Visible = False
        'LimpiarBusqueda
    End If
    oConexion.Close
    Set oConexion = Nothing
    mo_Apariencia.ConfigurarFilasBiColores Me.grdPacientesEncontrados, sighentidades.GrillaConFilasBicolor
ErrBusq:
End Sub

Sub DeudasPendientesDeAnterioresAtenciones(oConexion As Connection)
        'Deudas
        ms_MensajeError = mo_AdminFacturacion.DevuelveDeudaPacienteDeAntencionesAnteriores(ml_idPaciente, oConexion, mo_CuentasAtencion.idCuentaAtencion)
        If ms_MensajeError <> "" Then
           MsgBox "Tiene Deudas Pendientes por Pagar" & Chr(13) & Chr(13) & ms_MensajeError, vbInformation, Me.Caption
           '
           ucMensajeParpadeando1.Visible = True
           ucMensajeParpadeando1.MensajeDeTexto = "Deudas:  " & ms_MensajeError
           'debb-29/02/2016 (inicio)
           If mi_Opcion = sghAgregar And InStr(ms_MensajeError, "<FALLECIDO>") > 0 Then
               btnCancelar_Click
           End If
           'debb-29/02/2016 (fin)
        Else
           '
           ucMensajeParpadeando1.Visible = False
           ucMensajeParpadeando1.MensajeDeTexto = ""
        End If
        ms_MensajeError = ""

End Sub

Private Sub btnCancelar_Click()
           Me.Visible = False
End Sub


Private Sub btnDxReferencia_Click()
    BusquedaDx ""

End Sub

Private Sub btnImprimeFichaSIS_Click()
   If mi_Opcion <> sghAgregar Then
      CargaDatosAlObjetosDeDatos
   End If
   Dim oFua As New SIGHSis.clFUA
   oFua.idCuentaAtencion = mo_Atenciones.idCuentaAtencion
   oFua.lcNombrePc = mo_lcNombrePc
   oFua.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
   oFua.idUsuario = ml_idUsuario
   oFua.Opcion = mi_Opcion
   oFua.idServicio = mo_cmbServicioIngreso.BoundText
   oFua.MostrarFormulario
   Set oFua = Nothing

End Sub

Private Sub btnImprimePreCta_Click()
   If txtNroCuenta.Text <> "" Then
      ImprimePreCuenta
   End If
End Sub

Sub ImprimePreCuenta()
    
    Dim lcPaciente As String
    Dim lcMedico As String
    Dim lcCola As String
    Dim oRsTmp3 As New Recordset
    If mi_Opcion <> sghAgregar Then
       Me.UcPacienteDatosAloj1.CargarDatosAlObjetoDatos mo_Pacientes
    End If
    lcPaciente = Trim(mo_Pacientes.ApellidoPaterno) & " " & Trim(mo_Pacientes.ApellidoMaterno) & " " & Trim(mo_Pacientes.PrimerNombre)
    If mo_Pacientes.SegundoNombre <> "" Then
       lcPaciente = lcPaciente & " " & Trim(mo_Pacientes.SegundoNombre)
    End If
    If mo_Pacientes.TercerNombre <> "" Then
      lcPaciente = lcPaciente & " " & Trim(mo_Pacientes.TercerNombre)
    End If
    lcMedico = lblNombreMedico.Text
    lcCola = ""
    ImpresionPreCuenta Me.txtFechaIngreso.Text, Me.txtHoraIngreso.Text, lcPaciente, _
                                mo_Pacientes.NroHistoriaClinica, Me.cmbServicioIngreso.Text, _
                                lcMedico, "PACIENTE EXTERNO (" & Trim(Me.cmbFuenteFinanciamiento.Text) & ")", _
                                mo_Atenciones.idAtencion, "", mo_Atenciones.idCuentaAtencion, _
                                Me.cmbFuenteFinanciamiento.Text, lcCola, ml_idUsuario, "", _
                                mo_Pacientes.FichaFamiliar, mo_Pacientes.IdTipoNumeracion, wxParametro216, _
                                wxParametro306
    Set oRsTmp3 = Nothing
    Me.Visible = False
End Sub
Sub ImpresionPreCuenta(FechaIngreso As String, HoraIngreso As String, Paciente As String, NroHistoriaClinica As Long, _
                       Servicio As String, Medico As String, lcTipoServicio As String, lnIdAtencion As Long, _
                       ml_idOrden As String, ml_idCuentaAtencion As String, lcFormaPago As String, lcNroCola As String, _
                       lnIdUsuario As Long, lcServicioDelTarifario As String, lcFichaFamiliar As String, _
                       lnIdTipoNumeracion As Long, lcParametro216 As String, lcParametro306 As String, Optional lnIdMedico As Long)
        Dim lcBuscaParametro As New SIGHDatos.Parametros
        Dim rsReporte As New Recordset
        Dim mrs_Tmp As New Recordset
        Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
        Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
        Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
        Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
        Dim lcNombreHospital As String, lnIdprograma As Long, lnIdTurno As Long, lcTurno As String
        Dim lcProfesionMedico As String
        Dim lcRUC As String, lcDireccion As String, lcTlefono As String
        Dim lcDevuelveFechaServidor As New SIGHDatos.Parametros
        Dim lcImpresora As String, lcImpresoraActual As String
        
        '******** Crystal Report
'        Dim oRptClaseCry As New rCrystal
'        If lcBuscaParametro.SeleccionaFilaParametro(216) <> "1" Then
'           oRptClaseCry.DestinoReporte = sghPantalla
'        Else
'           oRptClaseCry.DestinoReporte = sghImpresora
'        End If
'        oRptClaseCry.TipoReporte = "ImpresionPreCuenta"
'        oRptClaseCry.lcTipoServicio = lcTipoServicio
'        oRptClaseCry.LcIdCuentaAtencion = ml_idCuentaAtencion
'        oRptClaseCry.lcFormaPago = "IAFA: " & lcFormaPago
'        oRptClaseCry.lcNroCola = lcNroCola
'        oRptClaseCry.idUsuario = lnIdUsuario
'        oRptClaseCry.lcServicioDelTarifario = Mid(lcServicioDelTarifario, InStr(lcServicioDelTarifario, "=") + 1, 100)
'        oRptClaseCry.idAtencion = lnIdAtencion
'        oRptClaseCry.Show vbModal
'        Set oRptClaseCry = Nothing
        '******** dataReport
        'Set rsReporte = mo_AdminReportes.ReporteAtencionesParaHistoriaClinica(lnIdAtencion)
        'If rsReporte.RecordCount = 0 Then
        '    MsgBox "No existen datos", vbInformation, "Reporte"
        'Else
            lcNombreHospital = lcBuscaParametro.SeleccionaFilaParametro(205)
            lcRUC = lcBuscaParametro.SeleccionaFilaParametro(339)
            lcDireccion = lcBuscaParametro.SeleccionaFilaParametro(206)
            lcTlefono = lcBuscaParametro.SeleccionaFilaParametro(207)
            lcImpresora = Trim(lcBuscaParametro.SeleccionaFilaParametro(336))
            '
            With mrs_Tmp
                .Fields.Append "FechaIngreso", adVarChar, 100, adFldIsNullable
                .Fields.Append "Turno", adVarChar, 100, adFldIsNullable
                .Fields.Append "Paciente", adVarChar, 100, adFldIsNullable
                .Fields.Append "Usuario", adVarChar, 100, adFldIsNullable
                .Fields.Append "NroHistoriaClinica", adVarChar, 100, adFldIsNullable
                .Fields.Append "Servicio", adVarChar, 100, adFldIsNullable
                .Fields.Append "Medico", adVarChar, 100, adFldIsNullable
                .Fields.Append "Especialidad", adVarChar, 100, adFldIsNullable
                .Fields.Append "Interconsulta", adVarChar, 100, adFldIsNullable
                .Fields.Append "ColaTipoS", adVarChar, 100, adFldIsNullable
                .Fields.Append "FechaHoraImp", adVarChar, 100, adFldIsNullable
                .LockType = adLockOptimistic
                .Open
            End With
            '
            lcProfesionMedico = ""
            If lnIdMedico > 0 Then
               lcProfesionMedico = mo_reglasComunes.TiposEmpleadosSeleccionarIdMedico(lnIdMedico)
               If lcProfesionMedico <> "" Then
                  lcProfesionMedico = lcProfesionMedico
               End If
            End If
            '
            lcTurno = ""
            Set rsReporte = mo_ReglasDeProgMedica.CitasSeleccionarXfechaMedico(FechaIngreso, lnIdMedico)
            rsReporte.Filter = "horaInicio='" & HoraIngreso & "'"
            If rsReporte.RecordCount > 0 Then
               lnIdprograma = rsReporte!IdProgramacion
               rsReporte.Close
               Set rsReporte = mo_ReglasDeProgMedica.ProgramacionMedicaSeleccionarXidentificador(lnIdprograma)
               If rsReporte.RecordCount > 0 Then
                  lnIdTurno = rsReporte.Fields!IdTurno
                  rsReporte.Close
                  Set rsReporte = mo_ReglasDeProgMedica.TurnosSeleccionarPorIdentificador(lnIdTurno)
                  If rsReporte.RecordCount > 0 Then
                     lcTurno = " (" & Trim(rsReporte!Descripcion) & ")"
                  End If
               End If
            End If
            rsReporte.Close
            '
            mrs_Tmp.AddNew
            If lcBuscaParametro.SeleccionaFilaParametro(281) = "x" Then  'solo el HRA
            '**** Programa: se agrego campo Turno que ya no esta concatenado con la fecha  (sale en la siguiente linea)
            '**** Programado por:Eder Yamill Palomino Espinoza
            '**** Fecha: 06102014
               
               mrs_Tmp.Fields!FechaIngreso = "Fecha: " & FechaIngreso & IIf(lcParametro306 = "S", "", lcTurno)
              
            Else
               mrs_Tmp.Fields!FechaIngreso = "Fecha: " & FechaIngreso & " Hr: " & HoraIngreso '& lcTurno
            End If
            mrs_Tmp.Fields!Paciente = "Paciente: " & Paciente
            mrs_Tmp.Fields!Usuario = "Usuario: " & mo_ReglasCaja.SeleccionaDatosCajero(lnIdUsuario, sghUsuario)
            'debb-09/09/2015 (inicio)
            If lcParametro306 = "S" Then
               mrs_Tmp.Fields!FechaHoraImp = "Fecha y Hora de Impresión: " & _
                                             lcDevuelveFechaServidor.RetornaFechaServidorSQL & " " & _
                                             lcDevuelveFechaServidor.RetornaHoraServidorSQLserverFormatoGalenhos
            Else
               mrs_Tmp.Fields!FechaHoraImp = "Fecha Impres: " & lcDevuelveFechaServidor.RetornaFechaServidorSQL & " " & _
                                             lcDevuelveFechaServidor.RetornaHoraServidorSQLserverFormatoGalenhos
            End If
            'debb-09/09/2015 (fin)
'            mrs_Tmp.Fields!FechaHoraImp = "Fecha y Hora de Impresión: " & lcDevuelveFechaServidor.RetornaFechaServidorSQL & " " & lcDevuelveFechaServidor.RetornaHoraServidorSQLserverFormatoGalenhos
            If lcBuscaParametro.SeleccionaFilaParametro(277) = "S" And Trim(lcFichaFamiliar) <> "" Then
               mrs_Tmp.Fields!NroHistoriaClinica = "Ficha Familiar: " & Trim(lcFichaFamiliar)
            Else
                If lcParametro306 = "S" Then
                    mrs_Tmp.Fields!NroHistoriaClinica = "N°Historia: " & Trim(Str(NroHistoriaClinica)) & " " & mo_ReglasArchivoClinico.HistoriaClinicaEsNueva(NroHistoriaClinica, lnIdTipoNumeracion)
                Else
                    mrs_Tmp.Fields!NroHistoriaClinica = "N°Historia: " & Trim(Str(NroHistoriaClinica)) & " " & mo_ReglasArchivoClinico.HistoriaClinicaEsNueva(NroHistoriaClinica, lnIdTipoNumeracion) & "     No Cuenta: " & ml_idCuentaAtencion
                End If
            End If
            mrs_Tmp.Fields!Servicio = "Serv: " & Servicio
            '**** Programa: se agrego campo Especialidad que ya no esta concatenado con medico (sale en la siguiente linea)
            '**** Programado por:Eder Yamill Palomino Espinoza
            '**** Fecha: 06102014
            'mrs_Tmp.Fields!Medico = Left("Médico: " & Trim(Medico) & lcProfesionMedico, 100)
            mrs_Tmp.Fields!Medico = lcProfesionMedico & " - " & Trim(Medico) '& lcProfesionMedico
            If lcNroCola = "" Then
                If IsNumeric(ml_idOrden) Then
                    mrs_Tmp.Fields!Interconsulta = "Ord. Pago: " & CStr(ml_idOrden)
                Else
                    mrs_Tmp.Fields!Interconsulta = ""
                End If
                mrs_Tmp.Fields!ColaTipoS = lcTipoServicio
            Else
                If lcParametro306 = "S" Then
                    mrs_Tmp.Fields!Interconsulta = "Interconsulta: Si (   )             No (  )"
                Else
                    mrs_Tmp.Fields!Interconsulta = "Interconsulta:           Si (   )             No (  )"
                End If
                mrs_Tmp.Fields!ColaTipoS = "Cupo:     " & lcNroCola
            End If
            mrs_Tmp.Update
            'Reporte
            On Error Resume Next
            '
            If Not (lcImpresora = "" Or lcImpresora = "0") Then
               lcImpresoraActual = Printer.DeviceName
               sighentidades.ImpresoraPredeterminada lcImpresora
            End If
            '
            If lcParametro306 = "S" Then
               '***************** Ticket a 1 copias en 1 pagina (Impresora: ticketera)**********************
                '**** Programa: se agrego Turno y Especialidad
                '**** Programado por:Eder Yamill Palomino Espinoza
                '**** Fecha: 06102014
                Set CePreCta.DataSource = mrs_Tmp
                CePreCta.Sections("Cabecera").Controls("lcHospital1").Caption = UCase(lcNombreHospital)
                CePreCta.Sections("Cabecera").Controls("DatosEESS").Caption = UCase("RUC: " & lcRUC) & vbCrLf & UCase(lcDireccion) & vbCrLf & "Telef: " & UCase(lcTlefono)
                CePreCta.Sections("Cabecera").Controls("NombreTicket").Caption = UCase("TICKET DE CITA")
'                CePreCta.Sections("Detalle").Controls("Turno").Caption = "Turno: " & lcTurno
                If lcTurno <> "" Then lcTurno = "Turno: " & lcTurno
                CePreCta.Sections("Detalle").Controls("Turno").Caption = lcTurno
                CePreCta.Sections("Detalle").Controls("lcNroHNroC").Caption = mrs_Tmp.Fields!NroHistoriaClinica
                CePreCta.Sections("Detalle").Controls("lcNroCuenta").Caption = "N°Cuenta: " & ml_idCuentaAtencion
                CePreCta.Sections("Detalle").Controls("Especialidad").Caption = "Especialidad: " & lcProfesionMedico

                CePreCta.Sections("Detalle").Controls("lcFormaPago").Caption = lcFormaPago
                CePreCta.Sections("Detalle").Controls("lcServicioDelTarifario").Caption = lcServicioDelTarifario

                CePreCta.Sections("PieInforme").Controls("lcUsuario").Caption = "Fecha: " & Format(lcDevuelveFechaServidor.RetornaFechaHoraServidorSQL, sighentidades.DevuelveFechaSoloFormato_DMY_HM) & "  " & "Usua: " & mo_ReglasCaja.SeleccionaDatosCajero(lnIdUsuario, sghUsuario)
                CePreCta.Sections("PieInforme").Controls("Terminal").Caption = "Terminal:  " & lcBuscaParametro.RetornaNombreDeServidor
                CePreCta.Sections("PieInforme").Controls("TextoVacio").Caption = lcBuscaParametro.SeleccionaFilaParametro(346)
                CePreCta.Sections("PieInforme").Controls("FechahoraImp").Caption = "Fecha y Hora de Impresión: " & lcDevuelveFechaServidor.RetornaFechaServidorSQL & " " & lcDevuelveFechaServidor.RetornaHoraServidorSQLserverFormatoGalenhos
                
                CePreCta.RightMargin = 100
                CePreCta.TopMargin = 100
                CePreCta.LeftMargin = 100
                CePreCta.BottomMargin = 100
                
                CePreCta.Orientation = rptOrientPortrait
                If lcParametro216 <> "1" Then
                   CePreCta.Show 1
                Else
                   CePreCta.PrintReport
                End If
            Else
                '***************** Ticket a 2 copias en 1 pagina (Impresora: EPSON)**********************
                Set CePreCuenta.DataSource = mrs_Tmp
                CePreCuenta.Sections("Detalle").Controls("lcHospital1").Caption = UCase(lcNombreHospital)
                CePreCuenta.Sections("Detalle").Controls("lcHospital2").Caption = UCase(lcNombreHospital)
                CePreCuenta.Sections("Detalle").Controls("lcFormaPago").Caption = lcFormaPago
                CePreCuenta.Sections("Detalle").Controls("lcFormaPago1").Caption = lcFormaPago
                CePreCuenta.Sections("Detalle").Controls("lcServicioDelTarifario").Caption = lcServicioDelTarifario
                CePreCuenta.Sections("Detalle").Controls("lcServicioDelTarifario1").Caption = lcServicioDelTarifario
                'debb-09/09/2015 (inicio)
                CePreCuenta.Sections("Detalle").Controls("mensaje").Caption = lcBuscaParametro.SeleccionaFilaParametro(346)
                CePreCuenta.Sections("Detalle").Controls("mensaje1").Caption = lcBuscaParametro.SeleccionaFilaParametro(346)
                'debb-09/09/2015 (fin)
                CePreCuenta.Orientation = rptOrientPortrait
                If lcParametro216 <> "1" Then
                   CePreCuenta.Show 1
                Else
                   CePreCuenta.PrintReport
                End If
            End If
            '
            If Not (lcImpresora = "" Or lcImpresora = "0") Then
               sighentidades.ImpresoraPredeterminada lcImpresoraActual
            End If
            '
        'End If
        'Set rsReporte = Nothing
        Set mrs_Tmp = Nothing
        Set mo_ReglasCaja = Nothing
        Set mo_reglasComunes = Nothing
        Set mo_ReglasArchivoClinico = Nothing
        Set mo_ReglasDeProgMedica = Nothing
End Sub


Private Sub btnLimpiar_Click()
    txtNroDNIBusqueda.Text = ""
    txtNroHistoriaBusqueda.Text = ""
    txtApellidoPaternoBusqueda.Text = ""
    txtApellidoMaternoBusqueda.Text = ""
    txtPrimerNombreBusqueda.Text = ""
    txtSegundoNombreBusqueda.Text = ""
    'UcSISafiliacion1.Limpiar
    On Error Resume Next
    txtNroDNIBusqueda.SetFocus

End Sub

Private Sub chkPacienteNuevo_Click()
    If chkPacienteNuevo.Value = 1 Then
       UcPacienteDatosAloj1.ActualizaDatosBasicos Me.txtApellidoPaternoBusqueda.Text, Me.txtApellidoMaternoBusqueda.Text, Me.txtPrimerNombreBusqueda.Text, Me.txtSegundoNombreBusqueda.Text, "00:00", 0
       If Val(lcBuscaParametro.SeleccionaFilaParametro(255)) = sghHistoriaDefinitivaManual Then
          UcPacienteDatosAloj1.SetFocusOnHistoria
       Else
          UcPacienteDatosAloj1.SetFocusOnApellidoPaterno
       End If
    End If
End Sub




Private Sub cmbFuenteFinanciamiento_Click(Area As Integer)
        Dim oRsBuscaPacientesSis As New Recordset
        Dim mo_SisConsumoWeb As New SIGHNegocios.SisConsumoWeb
        Dim oConexionExterna As New Connection
        Dim oRsTmp76 As New Recordset
        Set oRsFormaPago = mo_AdminFacturacion.TiposFinanciamientosTarifaSeleccionarPorPlan(Val(cmbFuenteFinanciamiento.BoundText))
        Set cmbFormaPago.RowSource = oRsFormaPago
        cmbFormaPago.ListField = "Descripcion"
        cmbFormaPago.BoundColumn = "idTipoFinanciamiento"
        mo_Formulario.HabilitarDeshabilitar Me.cmbFormaPago, True
        If oRsFormaPago.RecordCount = 1 Then
           cmbFormaPago.BoundText = oRsFormaPago.Fields!idTipoFinanciamiento
        ElseIf Val(cmbFuenteFinanciamiento.BoundText) = 5 Then
           cmbFormaPago.BoundText = "1"
        End If
        lnAfiliacionSIS4 = 0: lcSIScodigo = "": lcCodigoEstablecimientoAdscripcionSIS = ""
        If wxParametro302 = "S" Then
            If cmbFuenteFinanciamiento.Locked = False And Val(cmbFuenteFinanciamiento.BoundText) = sghFuenteFinanciamiento.sghFFSIS Then
               Dim lcDNI As String, lbPreguntar As Boolean
               
               If mo_ReglasSISgalenhos.PacienteBuscadoEnTablaGalenHosTieneAfiliacionSIS(lcDNI_busqueda, _
                                            UcPacienteDatosAloj1.DevuelveApaterno, UcPacienteDatosAloj1.DevuelveAmaterno, _
                                            UcPacienteDatosAloj1.DevuelvePnombre, UcPacienteDatosAloj1.DevuelveSnombre, _
                                            UcPacienteDatosAloj1.DevuelveSexo, UcPacienteDatosAloj1.DevuelveFechaNacimiento, _
                                            wxParametroJAMO, ldFechaActualServidor, lnAfiliacionSIS4, lcSIScodigo, False) = False Then
                                            
                   Set oRsBuscaPacientesSis = mo_SisConsumoWeb.WebServiceSISBuscarAfiliado(lcDNI_busqueda, "", _
                                                      "", "", "", "", wxParametro323)
                   If oRsBuscaPacientesSis.RecordCount > 0 Then
                      If mo_ReglasSISgalenhos.Sis_ValidaSiEsAfiliadoActualDelSIS(oRsBuscaPacientesSis, ldFechaActualServidor) = True Then
                            lnAfiliacionSIS4 = oRsBuscaPacientesSis!idSiasis
                            lcSIScodigo = oRsBuscaPacientesSis!Codigo
                            lcCodigoEstablecimientoAdscripcionSIS = IIf(IsNull(oRsBuscaPacientesSis!CodigoEstablAdscripcion), "", oRsBuscaPacientesSis!CodigoEstablAdscripcion)
                       Else
                            MsgBox "Ese Paciente ya venció"
                            cmbFuenteFinanciamiento.BoundText = ""
                            cmbFormaPago.BoundText = ""
                       End If
                   Else
                       MsgBox "No se encontró Paciente en tabla de FILIACIONES DEL SIS (sigh_externa)" & Chr(13) & Chr(13) & _
                              "Los Apellidos, Nombres, Sexo, F.Nacimiento (SisGalenPlus) deben ser iguales en la tabla de FILIACIONES (SIS)" & Chr(13) & _
                              "<<Haga una CITA y la ANULA, luego usa esta opción, en caso usa la WEB de Afiliados SIS>>", vbInformation, Me.Caption
                       
                       cmbFuenteFinanciamiento.BoundText = ""
                       cmbFormaPago.BoundText = ""
                    End If
               Else
                    oConexionExterna.CommandTimeout = 300
                    oConexionExterna.CursorLocation = adUseClient
                    oConexionExterna.Open wxParametroJAMO
                    Set oRsTmp76 = mo_ReglasSISgalenhos.SisFiliacionesSeleccionarPorIdSiaSis(lnAfiliacionSIS4, lcSIScodigo, oConexionExterna)
                    If oRsTmp76.RecordCount > 0 Then
                       lcCodigoEstablecimientoAdscripcionSIS = IIf(IsNull(oRsTmp76!CodigoEstablAdscripcion), "", oRsTmp76!CodigoEstablAdscripcion)
                    End If
                    oRsTmp76.Close
                    oConexionExterna.Close
               End If
               CargarAutomaticamenteEstablecimientoReferenciaSIS
               On Error Resume Next
               btnAceptar.SetFocus
            End If
        End If
        Set oRsBuscaPacientesSis = Nothing
        Set mo_SisConsumoWeb = Nothing
        Set oConexionExterna = Nothing
        Set oRsTmp76 = Nothing
End Sub

Private Sub cmbFuenteFinanciamiento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbFuenteFinanciamiento
    AdministrarKeyPreview KeyCode

End Sub

Private Sub cmbServicioIngreso_GotFocus()
   cmbServicioIngreso.SetFocus
End Sub

Private Sub cmbServicioIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbServicioIngreso
    AdministrarKeyPreview KeyCode

End Sub

Private Sub cmdSinApellidoMaterno_Click()
    txtApellidoMaternoBusqueda.Text = wxSinApellido
End Sub

Private Sub cmdSinApellidoPaterno_Click()
    txtApellidoPaternoBusqueda.Text = wxSinApellido
End Sub

Private Sub Form_Initialize()
    Set mo_cmbServicioIngreso.MiComboBox = cmbServicioIngreso
    Set mo_cmbIdTipoEdad.MiComboBox = cmbIdTipoEdad
    Set mo_cmbIdTipoReferenciaOrigen.MiComboBox = cmbIdTipoReferenciaOrigen
End Sub

Private Sub Form_Load()
    mo_Formulario.HabilitarDeshabilitar txtIdMedicoIngreso, False
    mo_Formulario.HabilitarDeshabilitar txtNroCuenta, False
    mo_Formulario.HabilitarDeshabilitar Me.txtEdadEnDias, False
    mo_Formulario.HabilitarDeshabilitar cmbIdTipoEdad, False
    mo_Formulario.HabilitarDeshabilitar txtIdEstablecimientoOrigen, False
    mo_Formulario.HabilitarDeshabilitar txtNombreOrigenReferencia, False
    mo_Formulario.HabilitarDeshabilitar txtDxReferencia, False
    mo_Formulario.HabilitarDeshabilitar lblDxReferencia1, False
    CargaDataCombos
    UcPacienteDatosAloj1.IdTipoGenHistoriaClinica = lcBuscaParametro.SeleccionaFilaParametro(255)
    UcPacienteDatosAloj1.Opcion = mi_Opcion
    UcPacienteDatosAloj1.inicializar
    UcPacienteDatosAloj1.HabilitaTipoHistoria True
    lnEspecialidadServicio = 0
    Me.txtFechaIngreso = Date
    Me.txtHoraIngreso = Format(Now, sighentidades.DevuelveHoraSoloFormato_HM)
    Select Case mi_Opcion
    Case sghAgregar
        Me.Caption = "Agregar Paciente Externo con Seguro"
    Case sghModificar
        Me.Caption = "Modificar Paciente Externo con Seguro"
    Case sghConsultar
        Me.Caption = "Consultar Paciente Externo con Seguro"
    Case sghEliminar
        Me.Caption = "Eliminar Paciente Externo con Seguro"
    End Select
    CargarDatosAlFormulario
    '
     
     InicilizarParametros
     If mi_Opcion = sghAgregar Then
        lnAfiliacionSIS4 = 0
        lcSIScodigo = ""
     End If
     '
End Sub

Sub CargarDatosAlFormulario()
     Select Case mi_Opcion
     Case sghAgregar
     Case sghModificar
         CargarDatosALosControles
     Case sghConsultar
         CargarDatosALosControles
     Case sghEliminar
         CargarDatosALosControles
     End Select
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   oRsFuentesFinanciamiento.Close
   Set oRsFuentesFinanciamiento = Nothing
End Sub

Private Sub grdPacientesEncontrados_DblClick()
    Dim rsPaciente As Recordset
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    On Error Resume Next
    Set rsPaciente = Me.grdPacientesEncontrados.DataSource
    If mo_AdminAdmision.BuscaSiEstaHospitalizado(rsPaciente!idPaciente, oConexion, sghConsultaExterna) = True Then  'debb-05/12/2015
       Exit Sub
    End If
    Me.grdPacientesEncontrados.Visible = False
    chkPacienteNuevo.Value = 0
    ml_idPaciente = rsPaciente!idPaciente
    UcPacienteDatosAloj1.idPaciente = rsPaciente!idPaciente
    UcPacienteDatosAloj1.CargarDatosDePacienteALosControles
    CalculaEdadEnLaAtencion
    DeudasPendientesDeAnterioresAtenciones oConexion
    LimpiarBusqueda
    cmbServicioIngreso.SetFocus
    oConexion.Close
    Set oConexion = Nothing
End Sub

Private Sub grdPacientesEncontrados_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdPacientesEncontrados.Bands(0).Columns("IdPaciente").Hidden = True
    grdPacientesEncontrados.Bands(0).Columns("IdTipoNumeracion").Hidden = True
    grdPacientesEncontrados.Bands(0).Columns("TipoNumeracion").Hidden = True
    'grdPacientesEncontrados.Bands(0).Columns("FechaNacimiento").Hidden = True
    
    grdPacientesEncontrados.Bands(0).Columns("NroHistoriaClinica").Header.Caption = "NroHistoria"
    grdPacientesEncontrados.Bands(0).Columns("NroHistoriaClinica").Width = 1000
    
    grdPacientesEncontrados.Bands(0).Columns("ApellidoPaterno").Header.Caption = "Ap. Paterno"
    grdPacientesEncontrados.Bands(0).Columns("ApellidoPaterno").Width = 1200
    
    grdPacientesEncontrados.Bands(0).Columns("ApellidoMaterno").Header.Caption = "Ap. Materno"
    grdPacientesEncontrados.Bands(0).Columns("ApellidoMaterno").Width = 1200
    
    grdPacientesEncontrados.Bands(0).Columns("PrimerNombre").Header.Caption = "1er Nombre"
    grdPacientesEncontrados.Bands(0).Columns("PrimerNombre").Width = 1200

    grdPacientesEncontrados.Bands(0).Columns("SegundoNombre").Header.Caption = "2do Nombre"
    grdPacientesEncontrados.Bands(0).Columns("SegundoNombre").Width = 1000

    grdPacientesEncontrados.Bands(0).Columns("FechaNacimiento").Header.Caption = "Fecha Nac."
    grdPacientesEncontrados.Bands(0).Columns("FechaNacimiento").Width = 1000

    grdPacientesEncontrados.Bands(0).Columns("TipoNumeracion").Header.Caption = "Tipo Numeración"
    grdPacientesEncontrados.Bands(0).Columns("TipoNumeracion").Width = 1500
    grdPacientesEncontrados.Bands(0).Columns("TipoNumeracion").CellAppearance.TextAlign = ssAlignRight

    grdPacientesEncontrados.Bands(0).Columns("TipoServicio").Header.Caption = "Ult. Tipo Serv."
    grdPacientesEncontrados.Bands(0).Columns("TipoServicio").Width = 1000

    grdPacientesEncontrados.Bands(0).Columns("FechaIngreso").Header.Caption = "Ult.Fec.Ing"
    grdPacientesEncontrados.Bands(0).Columns("FechaIngreso").Width = 1000

    grdPacientesEncontrados.Bands(0).Columns("FechaEgreso").Header.Caption = "Ult.Fec.Egr."
    grdPacientesEncontrados.Bands(0).Columns("FechaEgreso").Width = 1000

    grdPacientesEncontrados.Bands(0).Columns("ServicioIngreso").Header.Caption = "Ult. Serv. Ing."
    grdPacientesEncontrados.Bands(0).Columns("ServicioIngreso").Width = 2700

End Sub

Private Sub grdPacientesEncontrados_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = 13 Then
        grdPacientesEncontrados_DblClick
    End If

End Sub












Private Sub txtApellidoMaternoBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtApellidoMaternoBusqueda

End Sub

Private Sub txtApellidoMaternoBusqueda_LostFocus()
   txtApellidoMaternoBusqueda.Text = mo_Teclado.CapitalizarNombres(txtApellidoMaternoBusqueda.Text)
   If Len(txtApellidoMaternoBusqueda.Text) > 0 Then
      btnBuscarPaciente_Click
   End If

End Sub

Private Sub txtApellidoPaternoBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtApellidoPaternoBusqueda
End Sub



Private Sub txtApellidoPaternoBusqueda_LostFocus()
   txtApellidoPaternoBusqueda.Text = mo_Teclado.CapitalizarNombres(txtApellidoPaternoBusqueda.Text)
   If Len(txtApellidoPaternoBusqueda.Text) > 0 Then
      btnBuscarPaciente_Click
   End If
End Sub



Private Sub txtFechaIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaIngreso
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtFechaIngreso_LostFocus()
    If Not esfecha(txtFechaIngreso.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
        txtFechaIngreso.Text = sighentidades.FECHA_VACIA_DMY_HM
        Exit Sub
    ElseIf CDate(txtFechaIngreso.Text > ldFechaActualServidor) Then
        MsgBox "La fecha ingresada no puede ser mayo a la de HOY", vbInformation, Me.Caption
        txtFechaIngreso.Text = sighentidades.FECHA_VACIA_DMY_HM
        Exit Sub
    End If
    CalculaEdadEnLaAtencion
End Sub

Sub CalculaEdadEnLaAtencion()
    On Error Resume Next
    Me.txtEdadEnDias.Text = ""
    Dim oEdad As Edad
    oEdad = sighentidades.CalcularEdad(CDate(Me.UcPacienteDatosAloj1.FechaNacimiento), CDate(txtFechaIngreso.Text))
    Me.txtEdadEnDias.Text = oEdad.Edad
    mo_cmbIdTipoEdad.BoundText = oEdad.TipoEdad

End Sub



Private Sub txtHoraIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaIngreso
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtHoraIngreso_LostFocus()
 If Not sighentidades.ValidaHora(txtHoraIngreso) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
             txtHoraIngreso = sighentidades.HORA_VACIA_HM
        End If
End Sub

Private Sub txtNroDNIBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
       mo_Teclado.RealizarNavegacion KeyCode, txtNroDNIBusqueda
End Sub


Private Sub txtNroDNIBusqueda_LostFocus()
   If Len(txtNroDNIBusqueda.Text) > 0 Then
      
      btnBuscarPaciente_Click
   End If

End Sub

Private Sub txtNroHistoriaBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNroHistoriaBusqueda

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

Private Sub txtPrimerNombreBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtPrimerNombreBusqueda

End Sub



Private Sub txtPrimerNombreBusqueda_LostFocus()
   txtPrimerNombreBusqueda.Text = mo_Teclado.CapitalizarNombres(txtPrimerNombreBusqueda.Text)
   'mo_Formulario.MarcarComoVacio txtPrimerNombreBusqueda
   If Len(txtPrimerNombreBusqueda.Text) > 0 Then
      btnBuscarPaciente_Click
   End If

End Sub

Private Sub txtSegundoNombreBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
       mo_Teclado.RealizarNavegacion KeyCode, txtSegundoNombreBusqueda
End Sub

Private Sub txtSegundoNombreBusqueda_LostFocus()
    txtSegundoNombreBusqueda.Text = mo_Teclado.CapitalizarNombres(txtSegundoNombreBusqueda.Text)
   If Len(txtSegundoNombreBusqueda.Text) > 0 Then
      btnBuscarPaciente_Click
   End If

End Sub


Sub CargaDataCombos()
    mo_cmbServicioIngreso.BoundColumn = "IdServicio"
    mo_cmbServicioIngreso.ListField = "DservicioHosp"
    Set mo_cmbServicioIngreso.RowSource = mo_AdminAdmision.DevuelveServiciosQueSonPuntosCarga("(1,2,3,4,5,6,7)", sghFiltraSoloActivos, sghPorDescServicio)
    '
    mo_cmbIdTipoEdad.BoundColumn = "IdTipoEdad"
    mo_cmbIdTipoEdad.ListField = "DescripcionLarga"
    Set mo_cmbIdTipoEdad.RowSource = mo_AdminServiciosComunes.TiposEdadSeleccionarTodos
    mo_cmbIdTipoEdad.BoundText = "1"    'Default Años
    '
    Set oRsFuentesFinanciamiento = mo_AdminFacturacion.FuentesFinanciamientoSeleccionarSoloConSeguros
    Set cmbFuenteFinanciamiento.RowSource = oRsFuentesFinanciamiento
    cmbFuenteFinanciamiento.ListField = "Descripcion"
    cmbFuenteFinanciamiento.BoundColumn = "idFuenteFinanciamiento"
    '
    Set oRsFormaPago = mo_AdminFacturacion.TiposFinanciamientoSeleccionarSoloConPlan
    Set cmbFormaPago.RowSource = oRsFormaPago
    cmbFormaPago.ListField = "Descripcion"
    cmbFormaPago.BoundColumn = "idTipoFinanciamiento"
    mo_Formulario.HabilitarDeshabilitar Me.cmbFormaPago, False
    
    mo_cmbIdTipoReferenciaOrigen.BoundColumn = "IdTipoReferencia"
    mo_cmbIdTipoReferenciaOrigen.ListField = "DescripcionLarga"
    Set mo_cmbIdTipoReferenciaOrigen.RowSource = mo_AdminServiciosComunes.TiposReferenciaSeleccionarTodos
    
    Set cmbServicioReferenciaO.ListSource = mo_AdminServiciosComunes.SuSalud_upsSeleccionarTodos   'debb-21/06/2016
End Sub

Private Sub lblNombreMedico_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, lblNombreMedico
   AdministrarKeyPreview KeyCode

End Sub

Private Sub lblNombreMedico_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lbUltimaTeclaPulsoENTER = True
    Else
        lbUltimaTeclaPulsoENTER = False
    End If

End Sub

Private Sub lblNombreMedico_LostFocus()
        If lblNombreMedico.Locked = False And lbUltimaTeclaPulsoENTER = True And lnEspecialidadServicio > 0 Then
           lbUltimaTeclaPulsoENTER = False
           CompletarDatosDeMedico txtIdMedicoIngreso, lblNombreMedico, lnEspecialidadServicio, lblNombreMedico.Text
           On Error Resume Next
           txtFechaIngreso.SetFocus
        End If
End Sub

Sub CompletarDatosDeMedico(txtMedico As TextBox, lblNombreMedico As TextBox, lIdEspecialidad As Long, lcFiltraMedico As String)
'Dim oBusqueda As New MedicosBusqueda
Dim oBusqueda As New SIGHNegocios.BuscaMedicos
Dim oDoMedico As New DOMedico
Dim oDOEmpleado As New dOEmpleado
Dim oDOEspecialidades As New Collection
Dim oConexion As New Connection
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    oBusqueda.IdEspecialidad = lIdEspecialidad
    If mi_Opcion = sghAgregar Then
        oBusqueda.NombreMedico = lcFiltraMedico
    End If
    'oBusqueda.Show 1
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
       If mo_AdminProgramacion.MedicosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado, oDoMedico, oDOEmpleado, oDOEspecialidades, oConexion) Then
            txtMedico.Text = oDOEmpleado.CodigoPlanilla
            txtMedico.Tag = oDoMedico.idMedico
            lblNombreMedico = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
       End If
    End If
    Set oBusqueda = Nothing
    Set oDoMedico = Nothing
    Set oDOEmpleado = Nothing
    Set oDOEspecialidades = Nothing
    oConexion.Close
    Set oConexion = Nothing
End Sub


Private Sub cmbServicioIngreso_LostFocus()
   If mo_cmbServicioIngreso.BoundText <> "" Then
      Dim oRsTmp As New Recordset
      Set oRsTmp = mo_AdminFacturacion.ServiciosSeleccionarPorFiltro(" idServicio=" & mo_cmbServicioIngreso.BoundText, sghPorDescripcion)
      If oRsTmp.RecordCount > 0 Then
         lnEspecialidadServicio = oRsTmp.Fields!IdEspecialidad
      End If
      oRsTmp.Close
      Set oRsTmp = Nothing
   End If
   On Error Resume Next
   lblNombreMedico.SetFocus
End Sub


Private Sub UcPacienteDatosAloj1_SePresionoTeclaEspecial(KeyCode As Integer)
    On Error Resume Next
    Select Case KeyCode
    Case vbKeyReturn
         CalculaEdadEnLaAtencion
         Dim oConexion As New Connection
         oConexion.Open sighentidades.CadenaConexion
         oConexion.CursorLocation = adUseClient
         DeudasPendientesDeAnterioresAtenciones oConexion
         oConexion.Close
         Set oConexion = Nothing
         lbUltimaTeclaPulsoENTER = True
         cmbServicioIngreso.SetFocus
    Case Else
         AdministrarKeyPreview KeyCode
    End Select

End Sub


Sub AdministrarKeyPreview(KeyCode As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        btnCancelar_Click
    Case vbKeyF2
        'btnAceptar_Click
    End Select
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------
Function AgregarDatos() As Boolean
    AgregarDatos = mo_AdminAdmision.AdmisionPacienteExternoConSeguroAgregar(mo_CuentasAtencion, mo_Atenciones, mo_Pacientes, _
                                                                            mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, _
                                                                            mo_DoAtencionDatosAdicionales)
    ms_MensajeError = mo_AdminAdmision.MensajeError

    If wxParametro302 = "S" And mo_Atenciones.IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
       mo_ReglasSISgalenhos.SisFiliacionesActualizarAfiliadoDesdeWEB lcDNI_busqueda, "", "", _
                                        "", "", "", wxParametro323
    End If

End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------
Function ModificarDatos() As Boolean
    ModificarDatos = mo_AdminAdmision.AdmisionPacienteExternoConSeguroModificar(mo_CuentasAtencion, mo_Atenciones, _
                                                                        mo_Pacientes, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, _
                                                                        mo_DoAtencionDatosAdicionales)
    ms_MensajeError = mo_AdminAdmision.MensajeError
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------
Function EliminarDatos() As Boolean
    Dim oConexion As New Connection
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    ms_MensajeError = mo_AdminAdmision.VerificaSiTieneMovimientoFarmaciaOservicio(mo_CuentasAtencion.idCuentaAtencion, mo_Atenciones.idTipoServicio, oConexion)
    oConexion.Close
    Set oConexion = Nothing
    If ms_MensajeError = "" Then
        mo_CuentasAtencion.idEstado = 9 'anulado
        mo_Atenciones.IdEstadoAtencion = 0  'anulado
        EliminarDatos = mo_AdminAdmision.AdmisionPacienteExternoConSeguroAnular(mo_CuentasAtencion, mo_Atenciones, mo_Pacientes, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
        ms_MensajeError = mo_AdminAdmision.MensajeError
    Else
        MsgBox ms_MensajeError & Chr(13) & "La Anulación tendrá que realizarlo FACTURACION ", vbInformation, Me.Caption
    End If
End Function


Sub CargarDatosALosControles()
Dim lcEstadoAtencion As String
Dim oDoMedico As New DOMedico
Dim oDOEmpleado As New dOEmpleado
Dim oRsTmp As New Recordset
Dim oDOEspecialidades As New Collection
Dim oConexion As New Connection
        oConexion.Open sighentidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        fraBusqueda.Enabled = False
        '1do:   CARGAR DATOS DE LA ATENCION
        Set mo_Atenciones = mo_AdminAdmision.AtencionesSeleccionarPorId(Me.idAtencion, oConexion)
        If mo_Atenciones.idAtencion = 0 Then
            'El registro ha sido eliminado, pero no se hizo el refresh
             Exit Sub
        End If
        With mo_Atenciones
                mo_cmbServicioIngreso.BoundText = .IdServicioIngreso
                Me.txtIdMedicoIngreso.Tag = .IdMedicoIngreso
                Me.txtHoraIngreso.Text = IIf(.HoraIngreso = "", sighentidades.HORA_VACIA_HM, .HoraIngreso)
                Me.txtFechaIngreso.Text = IIf(.FechaIngreso = 0, sighentidades.FECHA_VACIA_DMY, .FechaIngreso)
                Me.txtEdadEnDias.Text = .Edad
                Me.txtEdadEnDias.Tag = .Edad
                mo_cmbIdTipoEdad.BoundText = .IdTipoEdad
                cmbIdTipoEdad.Tag = .IdTipoEdad
'
                If mo_AdminProgramacion.MedicosSeleccionarPorId(.IdMedicoIngreso, oDoMedico, oDOEmpleado, oDOEspecialidades, oConexion) Then
                    Me.txtIdMedicoIngreso = oDOEmpleado.CodigoPlanilla
                    Me.lblNombreMedico = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
                Else
                    Me.lblNombreMedico = ""
                End If
                cmbFuenteFinanciamiento.BoundText = .IdFuenteFinanciamiento
                Me.cmbFormaPago.BoundText = .IdFormaPago
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
        End With
        '
        Set oRsTmp = mo_AdminFacturacion.ServiciosSeleccionarPorFiltro("idServicio=" & mo_cmbServicioIngreso.BoundText, sghPorDescripcion)
        If oRsTmp.RecordCount > 0 Then
           lnEspecialidadServicio = oRsTmp.Fields!IdEspecialidad
        End If
        '
        Set mo_CuentasAtencion = mo_AdminFacturacion.CuentasAtencionSeleccionarPorId(mo_Atenciones.idCuentaAtencion, oConexion)
        lblEstadoCta.Caption = mo_ReglasFarmacia.DevuelveEstadoActualDeEstadoCuenta("idEstado=" & mo_CuentasAtencion.idEstado, oConexion)
        If mo_CuentasAtencion.idEstado <> 1 And mo_CuentasAtencion.idEstado <> 12 Then
            btnAceptar.Enabled = False
        End If
        txtNroCuenta.Text = mo_CuentasAtencion.idCuentaAtencion
        '3to:   CARGAR DATOS DEL PACIENTE
        UcPacienteDatosAloj1.idPaciente = mo_Atenciones.idPaciente
        UcPacienteDatosAloj1.CargarDatosDePacienteALosControles
        '
        DeudasPendientesDeAnterioresAtenciones oConexion
        '
        UcPacienteDatosAloj1.CargarDatosAlObjetoDatos mo_Pacientes
        Me.Caption = Trim(Me.Caption) & "                HC: " & Trim(mo_Pacientes.NroHistoriaClinica) & " " & Trim(mo_Pacientes.ApellidoPaterno) & " " & Trim(mo_Pacientes.ApellidoMaterno) & " " & Trim(mo_Pacientes.PrimerNombre) & "     (Estado: " & lcEstadoAtencion & ")"
        'Ya tuvo movimientos(Farmacia/servicios), no podrá cambiar de plan
        If mi_Opcion = sghModificar Then
            ms_MensajeError = mo_AdminAdmision.VerificaSiTieneMovimientoFarmaciaOservicio(mo_CuentasAtencion.idCuentaAtencion, mo_Atenciones.idTipoServicio, oConexion)
            If ms_MensajeError <> "" Then
               mo_Formulario.HabilitarDeshabilitar Me.cmbFuenteFinanciamiento, False
               Me.ucMensajeParpadeando1.MensajeDeTexto = ms_MensajeError
               Me.ucMensajeParpadeando1.Visible = True
            End If
            ms_MensajeError = ""
        End If
        '
        Set mo_DoAtencionDatosAdicionales = mo_AdminAdmision.AtencionesDatosAdicionalesSeleccionarPorId(mo_Atenciones.idAtencion, oConexion)
        If Not (mo_DoAtencionDatosAdicionales Is Nothing) Then
            With mo_DoAtencionDatosAdicionales
                 .idAtencion = mo_Atenciones.idAtencion
                 lnAfiliacionSIS4 = .idSiasis
                 lcSIScodigo = .SisCodigo
                 '.FuaCodigoPrestacion
                'debb-21/06/2016 (inicio)
                mo_cmbIdTipoReferenciaOrigen.BoundText = .IdTipoReferenciaOrigen
                CompletarDatosDelEstablecimientoEnElLoad .IdEstablecimientoOrigen, .IdEstablecimientoNoMinsaOrigen, txtIdEstablecimientoOrigen, txtNombreOrigenReferencia, .IdTipoReferenciaOrigen
                txtReferenciaO.Text = .NroReferenciaOrigen
                Dim lcDxCodigo As String, lcDx As String
                PVcomboBoxUbicaPosicion .referenciaOservicio, cmbServicioReferenciaO
                txtDxReferencia.Tag = .referenciaOidDiagnostico
                mo_AdminServiciosComunes.DiagnosticosSeleccionarPorIdDevuelveDescripcion Val(txtDxReferencia.Tag), _
                                                                                         oConexion, lcDxCodigo, lcDx
                txtDxReferencia.Text = lcDxCodigo
                lblDxReferencia1.Text = lcDx
            End With
        End If
        '
        If wxParametro302 = "S" And mo_Atenciones.IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
            btnImprimeFichaSIS.Visible = True
        End If
        '
        Set oRsTmp = Nothing
        Set oDoMedico = Nothing
        Set oDOEmpleado = Nothing
        Set oDOEspecialidades = Nothing
        oConexion.Close
        Set oConexion = Nothing
End Sub


Sub InicilizarParametros()
         wxParametro216 = lcBuscaParametro.SeleccionaFilaParametro(216)
         wxParametro280 = lcBuscaParametro.SeleccionaFilaParametro(280)
         wxParametro282 = lcBuscaParametro.SeleccionaFilaParametro(282)
         wxParametro302 = lcBuscaParametro.SeleccionaFilaParametro(302)
         wxParametro306 = lcBuscaParametro.SeleccionaFilaParametro(306)
         wxParametro323 = lcBuscaParametro.SeleccionaFilaParametro(323)
         wxParametro326 = lcBuscaParametro.SeleccionaFilaParametro(326)
         wxParametro336 = lcBuscaParametro.SeleccionaFilaParametro(336)
         wxParametroJAMO = lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
         ldFechaActualServidor = lcBuscaParametro.RetornaFechaServidorSQL
         wxParametroSIS = lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
End Sub


Public Sub CargarAutomaticamenteEstablecimientoReferenciaSIS() 'Frank 2808
    If wxParametro326 = "S" And wxParametro302 = "S" And mi_Opcion = sghAgregar And _
                            Val(cmbFuenteFinanciamiento.BoundText) = sghFuenteFinanciamiento.sghFFSIS Then
       Dim lcCodigoSis As String
       Dim lcEstablecimientoOrigen As String
       Dim DOEstablecimiento As New DOEstablecimiento
       Dim oRsEstabNoMINSA As Recordset
       Dim lnIdOrigenDelPacienteDesdeFUA As Long
       'lnIdOrigenDelPacienteDesdeFUA = mo_AdminAdmision.TiposOrigenAtencionDevuelveIdSis(Val(mo_cmbIdViasAdmision.BoundText))
       
       If Val(wxParametro280) <> Val(lcCodigoEstablecimientoAdscripcionSIS) Then
          If wxParametro282 <> "S" Then 'Hospital
        '       If Not (lnIdOrigenDelPacienteDesdeFUA = "4" Or lnIdOrigenDelPacienteDesdeFUA = "6") Then 'Referido CE, ContraReferido
                    'mo_cmbIdViasAdmision.BoundText = "12"
                    If mo_AdminServiciosComunes.EstablecimientosSeleccionarPorCodigo(Right(lcCodigoEstablecimientoAdscripcionSIS, 5), DOEstablecimiento) = True Then
                        mo_cmbIdTipoReferenciaOrigen.BoundText = 1 'MINSA
                        txtIdEstablecimientoOrigen.Text = DOEstablecimiento.Codigo
                        txtIdEstablecimientoOrigen.Tag = DOEstablecimiento.IdEstablecimiento
                        txtNombreOrigenReferencia.Text = DOEstablecimiento.Nombre
                    Else
                        Set oRsEstabNoMINSA = mo_AdminServiciosComunes.EstablecimientosNoMinsaSeleccionarPorCodigo(Right(lcCodigoEstablecimientoAdscripcionSIS, 5))
                        If oRsEstabNoMINSA.RecordCount > 0 Then
                            oRsEstabNoMINSA.MoveFirst
                            mo_cmbIdTipoReferenciaOrigen.BoundText = 2 'NO MINSA
                            txtIdEstablecimientoOrigen.Text = oRsEstabNoMINSA.Fields!Codigo
                            txtIdEstablecimientoOrigen.Tag = oRsEstabNoMINSA.Fields!IdEstablecimientoNoMINSA
                            txtNombreOrigenReferencia.Text = oRsEstabNoMINSA.Fields!Nombre
                        End If
                        Set oRsEstabNoMINSA = Nothing
                    End If
           '    End If
          End If
       End If
       Set DOEstablecimiento = Nothing
    End If
End Sub
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
        Set oDODiagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
        If Not oDODiagnostico Is Nothing Then
            txtDxReferencia.Text = oDODiagnostico.CodigoCIE2004
            txtDxReferencia.Tag = oDODiagnostico.IdDiagnostico
            lblDxReferencia1.Text = oDODiagnostico.Descripcion
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
Sub CompletarDatosDeEstablecimiento(txtIdEstablecimiento As TextBox, lblNombreEstablecimiento As TextBox, lTipoReferencia As Long)
    
    If lTipoReferencia = 1 Then
        'Dim oBusqueda As New EstablecimientosBusqueda
        Dim oBusqueda As New SIGHNegocios.BuscaEstablecimientos
        Dim oDoEstablecimiento As New DOEstablecimiento
        'oBusqueda.Show 1
        oBusqueda.MostrarFormulario
        If oBusqueda.BotonPresionado = sghAceptar Then
        
            Set oDoEstablecimiento = mo_AdminServiciosComunes.EstablecimientosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
            If Not oDoEstablecimiento Is Nothing Then
                txtIdEstablecimiento.Tag = oDoEstablecimiento.IdEstablecimiento
                txtIdEstablecimiento.Text = oDoEstablecimiento.Codigo
                lblNombreEstablecimiento = oDoEstablecimiento.Nombre
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
            Set oDoEstablecimientoNM = mo_AdminServiciosComunes.EstablecimientosNoMinsaSeleccionarPorId(oBusquedaNM.IdRegistroSeleccionado)
            If Not oDoEstablecimientoNM Is Nothing Then
                txtIdEstablecimiento.Tag = oDoEstablecimientoNM.IdEstablecimientoNoMINSA
                txtIdEstablecimiento.Text = oDoEstablecimientoNM.IdEstablecimientoNoMINSA
                lblNombreEstablecimiento = oDoEstablecimientoNM.Nombre
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

Sub CompletarDatosDelEstablecimientoEnElLoad(lIdEstablecimiento As Long, lIdEstablecimientoNoMinsa As Long, txtIdEstablecimiento As TextBox, lblNombreEstablecimiento As TextBox, lTipoReferencia As Long)
                
    If lTipoReferencia = 1 Then
        Dim oDoEstablecimiento As New DOEstablecimiento
         Set oDoEstablecimiento = mo_AdminServiciosComunes.EstablecimientosSeleccionarPorId(lIdEstablecimiento)
         If Not oDoEstablecimiento Is Nothing Then
             txtIdEstablecimiento.Text = oDoEstablecimiento.Codigo
             txtIdEstablecimiento.Tag = oDoEstablecimiento.IdEstablecimiento
             lblNombreEstablecimiento = oDoEstablecimiento.Nombre
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
             lblNombreEstablecimiento = oDOEstablecimientoNoMinsa.Nombre
        Else
             txtIdEstablecimiento.Text = ""
             txtIdEstablecimiento.Tag = ""
             lblNombreEstablecimiento = ""
         End If
    End If

End Sub

Function PVcomboBoxDevuelveEleccion(cmbComboPV As PVComboBox) As String
           Dim oCampos() As String
           If cmbComboPV.ListIndex < 0 Then
               PVcomboBoxDevuelveEleccion = ""
           Else
               oCampos = Split(cmbComboPV.List(cmbComboPV.ListIndex), "|")
               PVcomboBoxDevuelveEleccion = oCampos(0)
           End If
End Function
