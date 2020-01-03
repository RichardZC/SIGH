VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form AdmisionEmergDetalle 
   Caption         =   "Form1"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12180
   Icon            =   "AdmisionEmergDetalle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   12180
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Height          =   870
      Left            =   75
      TabIndex        =   133
      Top             =   -15
      Width           =   2145
      Begin VB.CheckBox chkPacienteNuevo 
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
         Height          =   225
         Left            =   135
         TabIndex        =   0
         Top             =   165
         Width           =   1710
      End
      Begin VB.ComboBox cmbIdTipoGravedad 
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
         Left            =   120
         TabIndex        =   1
         Text            =   "(Gravedad)"
         Top             =   450
         Width           =   1890
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1245
      Left            =   60
      TabIndex        =   70
      Top             =   7860
      Width           =   12015
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "AdmisionEmergDetalle.frx":08CA
         DownPicture     =   "AdmisionEmergDetalle.frx":0D8E
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
         Left            =   6510
         Picture         =   "AdmisionEmergDetalle.frx":127A
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   330
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "AdmisionEmergDetalle.frx":1766
         DownPicture     =   "AdmisionEmergDetalle.frx":1BC6
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
         Left            =   4965
         Picture         =   "AdmisionEmergDetalle.frx":203B
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   330
         Width           =   1365
      End
      Begin VB.CommandButton btnImprimir 
         Caption         =   "Imprimir"
         Height          =   705
         Left            =   225
         Picture         =   "AdmisionEmergDetalle.frx":24B0
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   315
         Width           =   1245
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
      Height          =   885
      Left            =   2280
      TabIndex        =   67
      Top             =   -30
      Width           =   9795
      Begin VB.TextBox txtApellidoPaternoBusqueda 
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
         Left            =   1785
         TabIndex        =   3
         Top             =   435
         Width           =   1395
      End
      Begin VB.TextBox txtNroDNIBusqueda 
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
         Left            =   7710
         TabIndex        =   136
         Top             =   435
         Width           =   975
      End
      Begin VB.TextBox txtPrimerNombreBusqueda 
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
         Left            =   4860
         TabIndex        =   5
         Top             =   435
         Width           =   1380
      End
      Begin VB.CommandButton btnBuscarPaciente 
         Height          =   315
         Left            =   8760
         Picture         =   "AdmisionEmergDetalle.frx":2989
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   420
         Width           =   585
      End
      Begin VB.TextBox txtApellidoMaternoBusqueda 
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
         TabIndex        =   4
         Top             =   435
         Width           =   1560
      End
      Begin VB.TextBox txtSegundoNombreBusqueda 
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
         Left            =   6315
         TabIndex        =   6
         Top             =   435
         Width           =   1335
      End
      Begin PVCOMBOLibCtl.PVComboBox cmbNroHistoriaBusqueda 
         Height          =   315
         Left            =   135
         TabIndex        =   2
         Top             =   435
         Width           =   1590
         _Version        =   524288
         _cx             =   2805
         _cy             =   556
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
         PrimaryColumn   =   6
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
         ColumnHeaderStyle=   0
         VerticalGridLines=   -1  'True
         HorizontalGridLines=   -1  'True
         ColumnResize    =   0   'False
         ItemLabelResize =   0   'False
         AllowDBAutoConfig=   0   'False
         GridLineColor   =   13421772
         List            =   ""
         NullString      =   ""
         DropShadow      =   -1  'True
         Text            =   ""
         SortOnColumnHeaderClick=   0   'False
         DropEffect      =   1
         ColumnCount     =   8
         Column0.Heading =   "IdPaciente"
         Column0.Width   =   40
         Column0.Alignment=   0
         Column0.Hidden  =   -1  'True
         Column0.Name    =   "IdPaciente"
         Column0.Format  =   ""
         Column0.Bound   =   -1  'True
         Column0.Locked  =   0   'False
         Column0.HeaderAlignment=   0
         Column1.Heading =   "IdTipoGeneracion"
         Column1.Width   =   40
         Column1.Alignment=   0
         Column1.Hidden  =   -1  'True
         Column1.Name    =   "IdTipoGeneracion"
         Column1.Format  =   ""
         Column1.Bound   =   -1  'True
         Column1.Locked  =   0   'False
         Column1.HeaderAlignment=   0
         Column2.Heading =   "Ap. paterno"
         Column2.Width   =   50
         Column2.Alignment=   0
         Column2.Hidden  =   0   'False
         Column2.Name    =   "ApellidoPaterno"
         Column2.Format  =   ""
         Column2.Bound   =   -1  'True
         Column2.Locked  =   0   'False
         Column2.HeaderAlignment=   0
         Column3.Heading =   "Ap. materno"
         Column3.Width   =   50
         Column3.Alignment=   0
         Column3.Hidden  =   0   'False
         Column3.Name    =   "ApellidoMaterno"
         Column3.Format  =   ""
         Column3.Bound   =   -1  'True
         Column3.Locked  =   0   'False
         Column3.HeaderAlignment=   0
         Column4.Heading =   "1er Nombre"
         Column4.Width   =   50
         Column4.Alignment=   0
         Column4.Hidden  =   0   'False
         Column4.Name    =   "PrimerNombre"
         Column4.Format  =   ""
         Column4.Bound   =   -1  'True
         Column4.Locked  =   0   'False
         Column4.HeaderAlignment=   0
         Column5.Heading =   "2do Nombre"
         Column5.Width   =   50
         Column5.Alignment=   0
         Column5.Hidden  =   0   'False
         Column5.Name    =   "SegundoNombre"
         Column5.Format  =   ""
         Column5.Bound   =   -1  'True
         Column5.Locked  =   0   'False
         Column5.HeaderAlignment=   0
         Column6.Heading =   "Nro Historia Clinica"
         Column6.Width   =   40
         Column6.Alignment=   0
         Column6.Hidden  =   0   'False
         Column6.Name    =   "NroHistoriaClinica"
         Column6.Format  =   "0"
         Column6.Bound   =   -1  'True
         Column6.Locked  =   0   'False
         Column6.HeaderAlignment=   0
         Column7.Heading =   "Tipo Historia"
         Column7.Width   =   40
         Column7.Alignment=   0
         Column7.Hidden  =   0   'False
         Column7.Name    =   "TipoGeneracion"
         Column7.Format  =   ""
         Column7.Bound   =   -1  'True
         Column7.Locked  =   0   'False
         Column7.HeaderAlignment=   0
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
      Begin VB.Label Label50 
         Caption         =   " Nº Historia clínica    Apellido paterno   Apellido materno      1er nombre      2do nombre         DNI"
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
         Left            =   165
         TabIndex        =   69
         Top             =   225
         Width           =   8445
      End
      Begin VB.Label Label51 
         Caption         =   "(&F6)"
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
         Left            =   9390
         TabIndex        =   68
         Top             =   465
         Width           =   345
      End
   End
   Begin TabDlg.SSTab tabAdmision 
      Height          =   6945
      Left            =   60
      TabIndex        =   64
      Top             =   930
      Width           =   12045
      _ExtentX        =   21246
      _ExtentY        =   12250
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
      TabCaption(0)   =   "Datos del paciente (F10)"
      TabPicture(0)   =   "AdmisionEmergDetalle.frx":2D63
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ucPacientesDetalle1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Datos de la atención (F11)"
      TabPicture(1)   =   "AdmisionEmergDetalle.frx":2D7F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblObservaciones"
      Tab(1).Control(1)=   "fraDatosSeguro"
      Tab(1).Control(2)=   "fraDatosCuenta"
      Tab(1).Control(3)=   "Frame2"
      Tab(1).Control(4)=   "Frame1"
      Tab(1).Control(5)=   "fraDatosReferenciaDestino"
      Tab(1).Control(6)=   "fraDatosReferenciaOrigen"
      Tab(1).Control(7)=   "Frame10"
      Tab(1).Control(8)=   "btnBuscarServicios"
      Tab(1).Control(9)=   "btnBuscarMedicos"
      Tab(1).Control(10)=   "btnBuscarMedicosEgreso"
      Tab(1).Control(11)=   "btnBuscarEstablecimiento"
      Tab(1).Control(12)=   "btnBuscarEstablecimientoDestino"
      Tab(1).Control(13)=   "txtObservacion"
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "Datos de la atención (F12)"
      TabPicture(2)   =   "AdmisionEmergDetalle.frx":2D9B
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSTab1"
      Tab(2).ControlCount=   1
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
         Height          =   525
         Left            =   -73170
         MultiLine       =   -1  'True
         TabIndex        =   32
         Top             =   6225
         Width           =   3855
      End
      Begin VB.CommandButton btnBuscarEstablecimientoDestino 
         Caption         =   "..."
         Height          =   315
         Left            =   -66075
         TabIndex        =   38
         Top             =   3120
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton btnBuscarEstablecimiento 
         Caption         =   "..."
         Height          =   315
         Left            =   -66075
         TabIndex        =   35
         Top             =   1080
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton btnBuscarMedicosEgreso 
         Caption         =   "..."
         Height          =   315
         Left            =   -72240
         TabIndex        =   20
         Top             =   3120
         Width           =   315
      End
      Begin VB.CommandButton btnBuscarMedicos 
         Caption         =   "..."
         Height          =   315
         Left            =   -72225
         TabIndex        =   14
         Top             =   1725
         Width           =   315
      End
      Begin VB.CommandButton btnBuscarServicios 
         Caption         =   "..."
         Height          =   315
         Left            =   -72210
         TabIndex        =   12
         Top             =   1350
         Width           =   315
      End
      Begin VB.Frame Frame10 
         Caption         =   "Condición del paciente"
         Height          =   705
         Left            =   -69015
         TabIndex        =   127
         Top             =   6240
         Visible         =   0   'False
         Width           =   11685
         Begin VB.ComboBox cmbIdCondicionEnElEstablecimiento 
            Height          =   315
            Left            =   7665
            TabIndex        =   135
            Top             =   240
            Width           =   3825
         End
         Begin VB.ComboBox cmbIdCondicionEnElServicio 
            Height          =   315
            Left            =   1665
            TabIndex        =   134
            Top             =   285
            Width           =   3825
         End
         Begin VB.Label Label5 
            Caption         =   "En el establecimiento"
            Height          =   285
            Left            =   5940
            TabIndex        =   129
            Top             =   300
            Width           =   2265
         End
         Begin VB.Label Label1 
            Caption         =   "En el servicio"
            Height          =   285
            Left            =   300
            TabIndex        =   128
            Top             =   300
            Width           =   1785
         End
      End
      Begin VB.Frame fraDatosReferenciaOrigen 
         Caption         =   "Datos de origen de  referencia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   -68985
         TabIndex        =   81
         Top             =   375
         Visible         =   0   'False
         Width           =   5865
         Begin VB.ComboBox cmbIdTipoReferenciaOrigen 
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
            Left            =   1830
            TabIndex        =   33
            Top             =   330
            Width           =   3870
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
            Left            =   1830
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   1050
            Width           =   3885
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
            Left            =   1830
            TabIndex        =   34
            Top             =   690
            Width           =   1000
         End
         Begin VB.Label lblIdTipoReferenciaOrigen 
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
            Height          =   315
            Left            =   135
            TabIndex        =   84
            Top             =   405
            Width           =   2175
         End
         Begin VB.Label lblIdEstablecimientoOrigen 
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
            Height          =   240
            Left            =   150
            TabIndex        =   83
            Top             =   750
            Width           =   1665
         End
      End
      Begin VB.Frame fraDatosReferenciaDestino 
         Caption         =   "Datos de destino de  referencia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -68970
         TabIndex        =   77
         Top             =   2460
         Visible         =   0   'False
         Width           =   5865
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
            Left            =   1815
            TabIndex        =   36
            Top             =   300
            Width           =   3855
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
            Left            =   1815
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   1020
            Width           =   3885
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
            Height          =   330
            Left            =   1815
            TabIndex        =   37
            Top             =   660
            Width           =   1000
         End
         Begin VB.Label lblIdTipoReferenciaDestino 
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
            Height          =   315
            Left            =   135
            TabIndex        =   80
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label lblIdEstablecimientoDestino 
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
            Height          =   225
            Left            =   150
            TabIndex        =   79
            Top             =   720
            Width           =   1665
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos de egreso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -74820
         TabIndex        =   71
         Top             =   2460
         Width           =   5745
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
            Left            =   1620
            TabIndex        =   24
            Top             =   1740
            Width           =   3885
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
            Left            =   1620
            TabIndex        =   23
            Top             =   1380
            Width           =   3885
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
            TabIndex        =   18
            Top             =   300
            Width           =   3915
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
            Left            =   2955
            TabIndex        =   132
            TabStop         =   0   'False
            Top             =   660
            Width           =   2580
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
            Left            =   1620
            TabIndex        =   19
            Top             =   660
            Width           =   885
         End
         Begin MSMask.MaskEdBox txtHoraEgreso 
            Height          =   315
            Left            =   3075
            TabIndex        =   22
            Top             =   1020
            Width           =   750
            _ExtentX        =   1323
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
            TabIndex        =   21
            Top             =   1020
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
         Begin VB.Label Label43 
            Caption         =   "Medico egreso"
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
            TabIndex        =   76
            Top             =   690
            Width           =   1335
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
            TabIndex        =   75
            Top             =   360
            Width           =   1155
         End
         Begin VB.Label Label15 
            Caption         =   "Fecha egreso"
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
            TabIndex        =   74
            Top             =   1050
            Width           =   1170
         End
         Begin VB.Label Label46 
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
            Height          =   315
            Left            =   120
            TabIndex        =   73
            Top             =   1410
            Width           =   1155
         End
         Begin VB.Label Label48 
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
            Height          =   285
            Left            =   120
            TabIndex        =   72
            Top             =   1770
            Width           =   1425
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos de ingreso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   -74820
         TabIndex        =   66
         Top             =   390
         Width           =   5745
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
            TabIndex        =   10
            Top             =   600
            Width           =   3930
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
            TabIndex        =   9
            Top             =   240
            Width           =   3930
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
            Left            =   3000
            TabIndex        =   131
            TabStop         =   0   'False
            Top             =   1320
            Width           =   2580
         End
         Begin VB.TextBox lblNombreServicio 
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
            Left            =   3000
            TabIndex        =   130
            TabStop         =   0   'False
            Top             =   960
            Width           =   2580
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
            Left            =   1650
            TabIndex        =   13
            Top             =   1320
            Width           =   885
         End
         Begin VB.TextBox txtIdServicioIngreso 
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
            TabIndex        =   11
            Top             =   960
            Width           =   885
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
            Left            =   4950
            TabIndex        =   17
            Top             =   1680
            Width           =   630
         End
         Begin MSMask.MaskEdBox txtHoraIngreso 
            Height          =   315
            Left            =   3105
            TabIndex        =   16
            Top             =   1680
            Width           =   750
            _ExtentX        =   1323
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
            Left            =   1650
            TabIndex        =   15
            Top             =   1680
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
         Begin VB.Label lblViaAdmision 
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
            Left            =   120
            TabIndex        =   85
            Top             =   615
            Width           =   1155
         End
         Begin VB.Label lblIdMedicoIngreso 
            Caption         =   "Medico ingreso"
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
            TabIndex        =   53
            Top             =   1335
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
            Left            =   90
            TabIndex        =   52
            Top             =   975
            Width           =   1395
         End
         Begin VB.Label lblIdTipoServicio 
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
            Left            =   105
            TabIndex        =   51
            Top             =   270
            Width           =   1395
         End
         Begin VB.Label lblEdadEnDias 
            Caption         =   "Edad (años)"
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
            Left            =   3945
            TabIndex        =   55
            Top             =   1710
            Width           =   1005
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
            TabIndex        =   54
            Top             =   1710
            Width           =   1230
         End
      End
      Begin VB.Frame fraDatosCuenta 
         Caption         =   "Datos de la cuenta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -74820
         TabIndex        =   65
         Top             =   4650
         Width           =   5745
         Begin VB.ComboBox cmbIdTipoFinanciamiento 
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
            Left            =   1635
            TabIndex        =   25
            Top             =   270
            Width           =   3885
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
            Left            =   3735
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   630
            Width           =   1110
         End
         Begin VB.TextBox txtHoraApertura 
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
            Left            =   4905
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   630
            Width           =   600
         End
         Begin VB.TextBox txtFechaCierre 
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
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   990
            Width           =   1110
         End
         Begin VB.TextBox txtHoraCierre 
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
            Left            =   2835
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   990
            Width           =   600
         End
         Begin VB.TextBox lblCuentaAtencion 
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
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   630
            Width           =   1110
         End
         Begin VB.CheckBox chkAnularAtencion 
            Alignment       =   1  'Right Justify
            Caption         =   "¿Anular atención?"
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
            Left            =   3675
            TabIndex        =   31
            Top             =   1050
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "Nro Cuenta"
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
            TabIndex        =   57
            Top             =   660
            Width           =   1065
         End
         Begin VB.Label Label6 
            Caption         =   "Tipo Financiam."
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
            Left            =   105
            TabIndex        =   56
            Top             =   330
            Width           =   1755
         End
         Begin VB.Label Label11 
            Caption         =   "Fecha apertura"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   2940
            TabIndex        =   58
            Top             =   570
            Width           =   1125
         End
         Begin VB.Label Label12 
            Caption         =   "Fecha Cierre"
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
            TabIndex        =   59
            Top             =   1050
            Width           =   1125
         End
      End
      Begin VB.Frame fraDatosSeguro 
         Caption         =   "Datos del seguro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         Left            =   -68970
         TabIndex        =   46
         Top             =   4680
         Width           =   5835
         Begin VB.ComboBox cmbIdPlan 
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
            Left            =   1845
            TabIndex        =   40
            Top             =   585
            Width           =   3825
         End
         Begin VB.ComboBox cmbIdFuenteFinanciamiento 
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
            Left            =   1830
            TabIndex        =   39
            Top             =   225
            Width           =   3855
         End
         Begin VB.TextBox txtNroAutorizacion 
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
            Left            =   1830
            TabIndex        =   41
            Top             =   960
            Width           =   1305
         End
         Begin VB.TextBox txtNroPlaca 
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
            Left            =   4335
            TabIndex        =   42
            Top             =   960
            Width           =   1305
         End
         Begin VB.Label Label7 
            Caption         =   "Fuente Financiam."
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
            TabIndex        =   60
            Top             =   300
            Width           =   1905
         End
         Begin VB.Label Label10 
            Caption         =   "Plan Cobertura"
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
            TabIndex        =   61
            Top             =   630
            Width           =   1665
         End
         Begin VB.Label Label13 
            Caption         =   "Nº Autorización"
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
            TabIndex        =   62
            Top             =   990
            Width           =   1545
         End
         Begin VB.Label lblNroPlaca 
            Caption         =   "Nº Placa"
            Height          =   255
            Left            =   3555
            TabIndex        =   63
            Top             =   990
            Width           =   885
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   6315
         Left            =   -74850
         TabIndex        =   50
         Top             =   450
         Width           =   11745
         _ExtentX        =   20717
         _ExtentY        =   11139
         _Version        =   393216
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
         TabPicture(0)   =   "AdmisionEmergDetalle.frx":2DB7
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "ucDiagnosticoDetalle1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Procedimientos"
         TabPicture(1)   =   "AdmisionEmergDetalle.frx":2DD3
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "ucProcedimientoDetalle1"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Exámenes"
         TabPicture(2)   =   "AdmisionEmergDetalle.frx":2DEF
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "ucExamenDetalle1"
         Tab(2).ControlCount=   1
         Begin VB.Frame Frame6 
            Caption         =   "Examenes"
            Height          =   1785
            Left            =   -74880
            TabIndex        =   104
            Top             =   390
            Width           =   11475
            Begin VB.TextBox Text6 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1680
               TabIndex        =   112
               Top             =   240
               Width           =   1005
            End
            Begin VB.CommandButton Command11 
               Caption         =   ".."
               Height          =   315
               Left            =   2760
               TabIndex        =   111
               Top             =   240
               Width           =   345
            End
            Begin VB.CommandButton Command10 
               Caption         =   "Agregar"
               Height          =   315
               Left            =   1650
               TabIndex        =   110
               Top             =   1320
               Width           =   1305
            End
            Begin VB.CommandButton Command7 
               Caption         =   "Eliminar"
               Height          =   315
               Left            =   3030
               TabIndex        =   109
               Top             =   1320
               Width           =   1305
            End
            Begin VB.CommandButton Command6 
               Caption         =   "..."
               Height          =   315
               Left            =   2760
               TabIndex        =   108
               Top             =   600
               Width           =   315
            End
            Begin VB.TextBox Text5 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1680
               TabIndex        =   107
               Top             =   600
               Width           =   975
            End
            Begin VB.CommandButton Command5 
               Caption         =   "..."
               Height          =   315
               Left            =   2760
               TabIndex        =   106
               Top             =   960
               Width           =   315
            End
            Begin VB.TextBox Text2 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1680
               TabIndex        =   105
               Top             =   960
               Width           =   975
            End
            Begin MSMask.MaskEdBox MaskEdBox3 
               Height          =   315
               Left            =   10740
               TabIndex        =   113
               Top             =   600
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Enabled         =   0   'False
               MaxLength       =   5
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox4 
               Height          =   315
               Left            =   9540
               TabIndex        =   114
               Top             =   600
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox5 
               Height          =   315
               Left            =   10740
               TabIndex        =   115
               Top             =   960
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Enabled         =   0   'False
               MaxLength       =   5
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox6 
               Height          =   315
               Left            =   9540
               TabIndex        =   116
               Top             =   960
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label45 
               Caption         =   "Fecha resultado"
               Height          =   315
               Left            =   7980
               TabIndex        =   124
               Top             =   990
               Width           =   1305
            End
            Begin VB.Label Label49 
               Caption         =   "Examen"
               Height          =   195
               Left            =   180
               TabIndex        =   123
               Top             =   300
               Width           =   1065
            End
            Begin VB.Label Label57 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   3150
               TabIndex        =   122
               Top             =   240
               Width           =   8175
            End
            Begin VB.Label Label58 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   3150
               TabIndex        =   121
               Top             =   600
               Width           =   4515
            End
            Begin VB.Label Label59 
               Caption         =   "Medico ordena"
               Height          =   315
               Left            =   150
               TabIndex        =   120
               Top             =   630
               Width           =   1335
            End
            Begin VB.Label Label60 
               Caption         =   "Fecha orden"
               Height          =   315
               Left            =   7980
               TabIndex        =   119
               Top             =   630
               Width           =   1305
            End
            Begin VB.Label Label61 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   3150
               TabIndex        =   118
               Top             =   960
               Width           =   4515
            End
            Begin VB.Label Label62 
               Caption         =   "Servicio ordena"
               Height          =   315
               Left            =   120
               TabIndex        =   117
               Top             =   990
               Width           =   1395
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Procedimientos"
            Height          =   1425
            Left            =   -74880
            TabIndex        =   86
            Top             =   360
            Width           =   11475
            Begin VB.TextBox Text4 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1680
               TabIndex        =   94
               Top             =   960
               Width           =   975
            End
            Begin VB.CommandButton Command9 
               Caption         =   "..."
               Height          =   315
               Left            =   2760
               TabIndex        =   93
               Top             =   960
               Width           =   315
            End
            Begin VB.TextBox Text3 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1680
               TabIndex        =   92
               Top             =   600
               Width           =   975
            End
            Begin VB.CommandButton Command8 
               Caption         =   "..."
               Height          =   315
               Left            =   2760
               TabIndex        =   91
               Top             =   600
               Width           =   315
            End
            Begin VB.CommandButton Command4 
               Caption         =   "Eliminar"
               Height          =   315
               Left            =   10050
               TabIndex        =   90
               Top             =   960
               Width           =   1305
            End
            Begin VB.CommandButton Command3 
               Caption         =   "Agregar"
               Height          =   315
               Left            =   8670
               TabIndex        =   89
               Top             =   960
               Width           =   1305
            End
            Begin VB.CommandButton Command2 
               Caption         =   ".."
               Height          =   315
               Left            =   2760
               TabIndex        =   88
               Top             =   240
               Width           =   345
            End
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1680
               TabIndex        =   87
               Top             =   240
               Width           =   1005
            End
            Begin MSMask.MaskEdBox MaskEdBox1 
               Height          =   315
               Left            =   10740
               TabIndex        =   95
               Top             =   600
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Enabled         =   0   'False
               MaxLength       =   5
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox2 
               Height          =   315
               Left            =   9540
               TabIndex        =   96
               Top             =   600
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label63 
               Caption         =   "Servicio realiza"
               Height          =   315
               Left            =   120
               TabIndex        =   103
               Top             =   990
               Width           =   1395
            End
            Begin VB.Label Label64 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   3150
               TabIndex        =   102
               Top             =   960
               Width           =   4515
            End
            Begin VB.Label Label65 
               Caption         =   "Fecha realización"
               Height          =   315
               Left            =   7980
               TabIndex        =   101
               Top             =   630
               Width           =   1305
            End
            Begin VB.Label Label66 
               Caption         =   "Medico realiza"
               Height          =   315
               Left            =   150
               TabIndex        =   100
               Top             =   630
               Width           =   1335
            End
            Begin VB.Label Label67 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   3150
               TabIndex        =   99
               Top             =   600
               Width           =   4515
            End
            Begin VB.Label Label68 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   3150
               TabIndex        =   98
               Top             =   240
               Width           =   8175
            End
            Begin VB.Label Label69 
               Caption         =   "Procedimiento"
               Height          =   195
               Left            =   180
               TabIndex        =   97
               Top             =   300
               Width           =   1065
            End
         End
         Begin UltraGrid.SSUltraGrid SSUltraGrid3 
            Height          =   2445
            Left            =   -74880
            TabIndex        =   125
            Top             =   2250
            Width           =   11475
            _ExtentX        =   20241
            _ExtentY        =   4313
            _Version        =   131072
            GridFlags       =   17040384
            LayoutFlags     =   67108864
            Caption         =   "Lista de procedimientos"
         End
         Begin UltraGrid.SSUltraGrid SSUltraGrid5 
            Height          =   2895
            Left            =   -74880
            TabIndex        =   126
            Top             =   1860
            Width           =   11475
            _ExtentX        =   20241
            _ExtentY        =   5106
            _Version        =   131072
            GridFlags       =   17040384
            LayoutFlags     =   67108864
            Caption         =   "Lista de procedimientos"
         End
         Begin Galenhos.ucDiagnosticoDetalle ucDiagnosticoDetalle1 
            Height          =   5745
            Left            =   120
            TabIndex        =   43
            Top             =   450
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   10134
         End
         Begin Galenhos.ucProcedimientoDetalle ucProcedimientoDetalle1 
            Height          =   5835
            Left            =   -74880
            TabIndex        =   44
            Top             =   360
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   10292
         End
         Begin Galenhos.ucExamenDetalle ucExamenDetalle1 
            Height          =   5835
            Left            =   -74880
            TabIndex        =   45
            Top             =   360
            Width           =   11475
            _ExtentX        =   20241
            _ExtentY        =   10292
         End
      End
      Begin Galenhos.ucPacientesDetalle ucPacientesDetalle1 
         Height          =   6435
         Left            =   150
         TabIndex        =   8
         Top             =   390
         Width           =   11715
         _ExtentX        =   20664
         _ExtentY        =   11351
      End
      Begin VB.Label lblObservaciones 
         Caption         =   "Datos del acompañante "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   -74700
         TabIndex        =   137
         Top             =   6225
         Width           =   1470
      End
   End
End
Attribute VB_Name = "AdmisionEmergDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase:
'        Autor: William Castro Grijalva
'        Fecha: 05/09/2004 01:20:06 p.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim mo_Teclado As New SIGHComun.Teclado
Dim mo_Formulario As New SIGHComun.Formulario
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim ml_IdUsuario As Long
Dim mb_ExistenDatos As Boolean
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminServiciosGeograficos As New SIGHNegocios.ReglasServGeograf
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_AdminFacturacion As New ReglasFacturacion
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_AdminProgramacion As New SIGHNegocios.ReglasDeProgMedica
Dim ml_TipoServicio As sghTipoServicio
Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
Dim mo_AdminHoteleria As New SIGHNegocios.ReglasHoteleria
Dim mrs_Diagnosticos As New ADODB.Recordset
Dim mo_Apariencia As New SIGHComun.GridInfragistic
Dim ml_TipoVistaForm As sghTipoVistaFormAtenciones
Dim ml_EstadoCuenta As Long
Dim mo_cmbIdTipoServicio As New SIGHComun.ListaDespleglable
Dim mo_cmbIdViasAdmision As New SIGHComun.ListaDespleglable
Dim mo_cmbIdEspecialidadMedico As New SIGHComun.ListaDespleglable
Dim mo_cmbIdServicio As New SIGHComun.ListaDespleglable
Dim mo_cmbIdDestinoAtencion As New SIGHComun.ListaDespleglable
Dim mo_cmbIdTipoFinanciamiento As New SIGHComun.ListaDespleglable
Dim mo_cmbIdCondicionEnElServicio As New SIGHComun.ListaDespleglable
Dim mo_cmbIdTipoReferenciaOrigen As New SIGHComun.ListaDespleglable
Dim mo_cmbIdTipoReferenciaDestino As New SIGHComun.ListaDespleglable
Dim mo_cmbIdFuenteFinanciamiento As New SIGHComun.ListaDespleglable
Dim mo_cmbIdPlan As New SIGHComun.ListaDespleglable
Dim mo_cmbIdCondicionEnElEstablecimiento As New SIGHComun.ListaDespleglable
Dim mo_cmbCondicionAlta As New SIGHComun.ListaDespleglable
Dim mo_cmbTipoAlta As New SIGHComun.ListaDespleglable
Dim mo_DoUbicacionPaciente As New doPaciente
'------------------------------------------------------------------------------------
'                               VARIABLES CUENTAS DE ATENCION
'------------------------------------------------------------------------------------
Dim mo_CuentasAtencion As New DOCuentaAtencion
Dim ml_IdCuentaAtencion As Long
Dim mo_cmbIdTipoGravedad As New SIGHComun.ListaDespleglable

'------------------------------------------------------------------------------------
'                               VARIABLE PARA LA ATENCION
'------------------------------------------------------------------------------------
Dim mo_Atenciones As New DOAtencion
Dim ml_IdAtencion As Long
Dim mo_Diagnosticos As New Collection
Dim mo_Procedimientos As New Collection
Dim mo_Examenes As New Collection

'------------------------------------------------------------------------------------
'                               VARIABLE PARA LA FILIACION
'------------------------------------------------------------------------------------
Dim mo_Pacientes  As New doPaciente
Dim ml_IdPaciente As Long
Dim ml_TipoGeneracionHistoria As sghTipoGeneracionDeNroHistoria
Dim mo_Historia As New DOHistoriaClinica

'------------------------------------------------------------------------------------
'                               VARIABLE PARA LA CITA
'------------------------------------------------------------------------------------
Dim ml_IdMedico As Long
Dim ms_NombreMedico  As String
Dim mo_Especialidad As New doEspecialidad
Dim mo_Paciente As New doPaciente
Dim ml_IdPrestamo As Long
Dim ml_IdEspecialidad As Long

Property Let IdMedico(lValue As Long)
   ml_IdMedico = lValue
End Property
Property Get IdMedico() As Long
   IdMedico = ml_IdMedico
End Property
Property Let IdPrestamo(lValue As Long)
   ml_IdPrestamo = lValue
End Property
Property Get IdPrestamo() As Long
   IdPrestamo = ml_IdPrestamo
End Property
Property Let IdPaciente(lValue As Long)
   ml_IdPaciente = lValue
End Property
Property Get IdPaciente() As Long
   IdPaciente = ml_IdPaciente
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
Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
End Property
Property Get IdUsuario() As Long
   IdUsuario = ml_IdUsuario
End Property
Property Let IdCuentaAtencion(lValue As Long)
   ml_IdCuentaAtencion = lValue
End Property
Property Get IdCuentaAtencion() As Long
   IdCuentaAtencion = ml_IdCuentaAtencion
End Property
Property Let IdAtencion(lValue As Long)
   ml_IdAtencion = lValue
End Property
Property Get IdAtencion() As Long
   IdAtencion = ml_IdAtencion
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
Property Let TipoVistaForm(lValue As sghTipoVistaFormAtenciones)
    ml_TipoVistaForm = lValue
End Property

Sub CargarComboBoxes()
Dim sSQL As String
Dim sMensaje As String

        'COMBO BOXES DE CUENTAS DE ATENCION
        mo_cmbIdTipoFinanciamiento.BoundColumn = "IdTipoFinanciamiento"
        mo_cmbIdTipoFinanciamiento.ListField = "DescripcionLarga"
        Set mo_cmbIdTipoFinanciamiento.RowSource = mo_AdminFacturacion.TiposFinanciamientoSeleccionarTodos
        sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError
       
        'COMBO BOXES DE ATENCION
        mo_cmbIdCondicionEnElServicio.BoundColumn = "IdTipoCondicionPaciente"
        mo_cmbIdCondicionEnElServicio.ListField = "DescripcionLarga"
        Set mo_cmbIdCondicionEnElServicio.RowSource = mo_AdminServiciosComunes.TiposCondicionPacienteSeleccionarTodos
        sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError
        
        mo_cmbIdCondicionEnElEstablecimiento.BoundColumn = "IdTipoCondicionPaciente"
        mo_cmbIdCondicionEnElEstablecimiento.ListField = "DescripcionLarga"
        Set mo_cmbIdCondicionEnElEstablecimiento.RowSource = mo_AdminServiciosComunes.TiposCondicionPacienteSeleccionarTodos
        sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError
        
        mo_cmbIdViasAdmision.BoundColumn = "IdOrigenAtencion"
        mo_cmbIdViasAdmision.ListField = "DescripcionLarga"
        Set mo_cmbIdViasAdmision.RowSource = mo_AdminAdmision.TiposOrigenAtencionSeleccionarViasDeConsultoriosExt
        sMensaje = sMensaje + mo_AdminAdmision.MensajeError
        
        mo_cmbIdDestinoAtencion.BoundColumn = "IdDestinoAtencion"
        mo_cmbIdDestinoAtencion.ListField = "DescripcionLarga"
        Set mo_cmbIdDestinoAtencion.RowSource = mo_AdminAdmision.TiposDestinoAtencionSeleccionarDestinosDeConsultorioEmergencia
        sMensaje = sMensaje + mo_AdminAdmision.MensajeError
        
       mo_cmbIdTipoReferenciaOrigen.BoundColumn = "IdTipoReferencia"
       mo_cmbIdTipoReferenciaOrigen.ListField = "DescripcionLarga"
       Set mo_cmbIdTipoReferenciaOrigen.RowSource = mo_AdminServiciosComunes.TiposReferenciaSeleccionarTodos
       sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError
       
       mo_cmbIdTipoReferenciaDestino.BoundColumn = "IdTipoReferencia"
       mo_cmbIdTipoReferenciaDestino.ListField = "DescripcionLarga"
       Set mo_cmbIdTipoReferenciaDestino.RowSource = mo_AdminServiciosComunes.TiposReferenciaSeleccionarTodos
       sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError
        
        mo_cmbIdTipoServicio.BoundColumn = "IdTipoServicio"
        mo_cmbIdTipoServicio.ListField = "DescripcionLarga"
        Set mo_cmbIdTipoServicio.RowSource = mo_AdminServiciosHosp.TiposServicioSeleccionarDeEmergencia
        mo_cmbIdTipoServicio.BoundText = "2"
        mo_Formulario.HabilitarDeshabilitar cmbIdTipoServicio, False
        
       mo_cmbCondicionAlta.BoundColumn = "IdCondicionAlta"
       mo_cmbCondicionAlta.ListField = "DescripcionLarga"
       Set mo_cmbCondicionAlta.RowSource = mo_AdminServiciosComunes.TiposCondicionAltaSeleccionarTodos
       sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError
        
       mo_cmbTipoAlta.BoundColumn = "IdTipoAlta"
       mo_cmbTipoAlta.ListField = "DescripcionLarga"
       Set mo_cmbTipoAlta.RowSource = mo_AdminServiciosComunes.TiposAltaSeleccionarTodos
       sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError
        
        Me.ucPacientesDetalle1.ConfigurarComboBoxes
        Me.ucDiagnosticoDetalle1.TipoDiagnostico = sghAtencionConsultaExterna
        Me.ucDiagnosticoDetalle1.ConfigurarComboBoxes
        
        Dim rsTiposGravedad As New ADODB.Recordset
        Set rsTiposGravedad = mo_AdminServiciosComunes.TipoGravedadAtencionSeleccionarTodos()
        mo_cmbIdTipoGravedad.CargarComboBoxDesdeRecordset cmbIdTipoGravedad, rsTiposGravedad, "IdTipoGravedad", "DescripcionLarga"
        
        '----------------------------------------------------------------------------------
       sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError
       If sMensaje <> "" Then
           MsgBox sMensaje, vbCritical, Me.Caption
       End If

End Sub


Private Sub btnBuscarEstablecimiento_Click()
    
    If Val(mo_cmbIdTipoReferenciaOrigen.BoundText) = 1 Then
        Dim oBusqueda As New EstablecimientosBusqueda
        Dim oDoEstablecimiento As New DOEstablecimiento
        oBusqueda.Show 1
        If oBusqueda.BotonPresionado = sghAceptar Then
            Set oDoEstablecimiento = mo_AdminServiciosComunes.EstablecimientosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
            If Not oDoEstablecimiento Is Nothing Then
                Me.txtIdEstablecimientoOrigen.Tag = oDoEstablecimiento.IdEstablecimiento
                Me.txtIdEstablecimientoOrigen.Text = oDoEstablecimiento.Codigo
                Me.txtNombreOrigenReferencia = oDoEstablecimiento.Nombre
            End If
        End If
    Else
        Dim oBusquedaNM As New EstablecimientosNoMinsaBusqueda
        Dim oDoEstablecimientoNM As New DOEstablecimientoNoMinsa
        oBusquedaNM.Show 1
        If oBusquedaNM.BotonPresionado = sghAceptar Then
            Set oDoEstablecimientoNM = mo_AdminServiciosComunes.EstablecimientosNoMinsaSeleccionarPorId(oBusquedaNM.IdRegistroSeleccionado)
            If Not oDoEstablecimientoNM Is Nothing Then
                Me.txtIdEstablecimientoOrigen.Tag = oDoEstablecimientoNM.IdEstablecimientoNoMinsa
                Me.txtIdEstablecimientoOrigen.Text = oDoEstablecimientoNM.IdEstablecimientoNoMinsa
                Me.txtNombreOrigenReferencia = oDoEstablecimientoNM.Nombre
            End If
        End If
    End If
End Sub

Private Sub btnBuscarEstablecimientoDestino_Click()
    If Val(mo_cmbIdTipoReferenciaDestino.BoundText) = 1 Then
        Dim oBusqueda As New EstablecimientosBusqueda
        Dim oDoEstablecimiento As New DOEstablecimiento
        oBusqueda.Show 1
        If oBusqueda.BotonPresionado = sghAceptar Then
            Set oDoEstablecimiento = mo_AdminServiciosComunes.EstablecimientosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
            If Not oDoEstablecimiento Is Nothing Then
                Me.txtIdEstablecimientoDestino.Tag = oDoEstablecimiento.IdEstablecimiento
                Me.txtIdEstablecimientoDestino.Text = oDoEstablecimiento.Codigo
                Me.txtNombreDestinoReferencia = oDoEstablecimiento.Nombre
            End If
        End If
    Else
        Dim oBusquedaNM As New EstablecimientosNoMinsaBusqueda
        Dim oDoEstablecimientoNM As New DOEstablecimientoNoMinsa
        oBusquedaNM.Show 1
        If oBusquedaNM.BotonPresionado = sghAceptar Then
            Set oDoEstablecimientoNM = mo_AdminServiciosComunes.EstablecimientosNoMinsaSeleccionarPorId(oBusquedaNM.IdRegistroSeleccionado)
            If Not oDoEstablecimientoNM Is Nothing Then
                Me.txtIdEstablecimientoDestino.Tag = oDoEstablecimientoNM.IdEstablecimientoNoMinsa
                Me.txtIdEstablecimientoDestino.Text = oDoEstablecimientoNM.IdEstablecimientoNoMinsa
                Me.txtNombreDestinoReferencia = oDoEstablecimientoNM.Nombre
            End If
        End If
    End If
End Sub

Private Sub btnBuscarMedicos_Click()
Dim oBusqueda As New MedicosBusqueda
Dim oDoMedico As New DOMedico
Dim oDOEmpleado As New DOEmpleado
Dim oDOEspecialidades As New Collection

    oBusqueda.Show 1
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        If mo_AdminProgramacion.MedicosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado, oDoMedico, oDOEmpleado, oDOEspecialidades) Then
            Me.txtIdMedicoIngreso.Text = oDOEmpleado.CodigoPlanilla
            Me.txtIdMedicoIngreso.Tag = oDoMedico.IdMedico
            Me.lblNombreMedico = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        End If
    End If

End Sub

Private Sub btnBuscarMedicosEgreso_Click()
Dim oBusqueda As New MedicosBusqueda
Dim oDoMedico As New DOMedico
Dim oDOEmpleado As New DOEmpleado
Dim oDOEspecialidades As New Collection

    oBusqueda.Show 1
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        If mo_AdminProgramacion.MedicosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado, oDoMedico, oDOEmpleado, oDOEspecialidades) Then
            Me.txtIdMedicoEgreso.Text = oDOEmpleado.CodigoPlanilla
            Me.txtIdMedicoEgreso.Tag = oDoMedico.IdMedico
            Me.lblNombreMedicoEgreso = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        End If
    End If

End Sub

Private Sub btnBuscarPaciente_Click()
Dim rsHistorias As New Recordset
Dim oDOPaciente As New doPaciente
    
    
    oDOPaciente.NroHistoriaClinica = Val(Me.cmbNroHistoriaBusqueda.Text)
    oDOPaciente.ApellidoPaterno = Me.txtApellidoPaternoBusqueda
    oDOPaciente.ApellidoMaterno = Me.txtApellidoMaternoBusqueda
    oDOPaciente.PrimerNombre = Me.txtPrimerNombreBusqueda
    oDOPaciente.SegundoNombre = Me.txtSegundoNombreBusqueda
    oDOPaciente.IdDocIdentidad = 1
    oDOPaciente.NroDocumento = Me.txtNroDNIBusqueda
    
    If (oDOPaciente.ApellidoPaterno + oDOPaciente.ApellidoMaterno + _
    oDOPaciente.PrimerNombre + oDOPaciente.SegundoNombre = "") And _
    (Val(Me.cmbNroHistoriaBusqueda.Text) = 0) And _
    (oDOPaciente.NroDocumento = "") Then
        MsgBox "Ingrese alguno de los valores de búsqueda", vbInformation, Me.Caption
        Exit Sub
    End If

    
    Set rsHistorias = mo_AdminAdmision.PacientesFiltrar(oDOPaciente)
    cmbNroHistoriaBusqueda.BoundColumn = ""
    Set cmbNroHistoriaBusqueda.ListSource = rsHistorias
    
    'Si hay una sola coincidencia
    If rsHistorias.RecordCount = 1 Then
        rsHistorias.MoveFirst
        chkPacienteNuevo.Value = 0
        
        Me.ucPacientesDetalle1.LimpiarDatosDePaciente
        
        Me.ucPacientesDetalle1.IdPaciente = rsHistorias!IdPaciente
        Me.ucPacientesDetalle1.NroHistoriaClinica = rsHistorias!NroHistoriaClinica
        Me.ucPacientesDetalle1.TipoGeneracionHistoriaClinica = rsHistorias!IdTipoGeneracion
        Me.IdPaciente = rsHistorias!IdPaciente
        Me.ucPacientesDetalle1.CargarDatosDePacienteALosControles
        
        Me.tabAdmision.tab = 0
    
    ElseIf rsHistorias.RecordCount > 1 Then
        cmbNroHistoriaBusqueda.ShowDropDown
        
    ElseIf rsHistorias.RecordCount = 0 Then
            
        LimpiarFormulario
        
        Me.ucPacientesDetalle1.TipoGeneracionHistoriaClinica = 0
        Me.ucPacientesDetalle1.NroHistoriaClinica = 0
        
        cmbNroHistoriaBusqueda.BoundText = ""
        txtApellidoMaternoBusqueda = ""
        txtPrimerNombreBusqueda = ""
        txtSegundoNombreBusqueda = ""
        txtApellidoPaternoBusqueda = ""
        txtNroDNIBusqueda = ""
        Me.tabAdmision.tab = 0
        
    End If
    rsHistorias.Close
End Sub

Private Sub btnBuscarServicios_Click()
    CompletarDatosDeServicio txtIdServicioIngreso, lblNombreServicio
End Sub

Private Sub btnImprimir_Click()
Dim oRptHistoriaConsultaExterna As New RptHistoriaConsEmerg

    If Me.IdAtencion = 0 Then
        MsgBox "De agregar la atención para poder imprimir", vbInformation, Me.Caption
        Exit Sub
    End If
    
    oRptHistoriaConsultaExterna.IdAtencion = Me.IdAtencion
    oRptHistoriaConsultaExterna.CrearReporteHistoriaClinicaConsultorioEmerg
    
End Sub

Private Sub chkAnularAtencion_Click()
    If Me.chkAnularAtencion = 1 Then
        Me.txtFechaCierre = Date
        Me.txtHoraCierre = Format(Now, "hh:mm")
    Else
        Me.txtFechaCierre = "__/__/____"
        Me.txtHoraCierre = "__:__"
    End If
End Sub

Private Sub chkPacienteNuevo_Click()
    
    If chkPacienteNuevo.Value = 1 Then
        If mi_Opcion = sghAgregar Then
            LimpiarFormulario
            
            Me.ucPacientesDetalle1.TipoGeneracionHistoriaClinica = 0
            Me.ucPacientesDetalle1.NroHistoriaClinica = 0
            
            cmbNroHistoriaBusqueda.BoundText = ""
            txtApellidoMaternoBusqueda = ""
            txtPrimerNombreBusqueda = ""
            txtSegundoNombreBusqueda = ""
            txtApellidoPaternoBusqueda = ""
            txtNroDNIBusqueda = ""
            Me.tabAdmision.tab = 0
            
            Me.ucPacientesDetalle1.ConfigurarValoresPorDefecto
        End If
    End If

    mo_Formulario.HabilitarDeshabilitar Me.cmbNroHistoriaBusqueda, Not (chkPacienteNuevo.Value = 1)
    mo_Formulario.HabilitarDeshabilitar Me.txtApellidoMaternoBusqueda, Not (chkPacienteNuevo.Value = 1)
    mo_Formulario.HabilitarDeshabilitar Me.txtPrimerNombreBusqueda, Not (chkPacienteNuevo.Value = 1)
    mo_Formulario.HabilitarDeshabilitar Me.txtSegundoNombreBusqueda, Not (chkPacienteNuevo.Value = 1)
    mo_Formulario.HabilitarDeshabilitar Me.txtApellidoPaternoBusqueda, Not (chkPacienteNuevo.Value = 1)
    mo_Formulario.HabilitarDeshabilitar Me.txtNroDNIBusqueda, Not (chkPacienteNuevo.Value = 1)


End Sub

Private Sub cmbNroHistoriaBusqueda_Click()
Dim oCampos() As String

    oCampos = Split(cmbNroHistoriaBusqueda.List(cmbNroHistoriaBusqueda.ListIndex), "|")
    
    If Val(oCampos(0)) <> 0 Then
        Me.ucPacientesDetalle1.LimpiarDatosDePaciente
        Me.ucPacientesDetalle1.TipoGeneracionHistoriaClinica = oCampos(1)
        Me.ucPacientesDetalle1.NroHistoriaClinica = Val(oCampos(6))
        Me.ucPacientesDetalle1.IdPaciente = oCampos(0)
        Me.IdPaciente = oCampos(0)
        Me.ucPacientesDetalle1.CargarDatosDePacienteALosControles
        chkPacienteNuevo.Value = 0
        Me.tabAdmision.tab = 0
    End If

End Sub

Private Sub cmbNroHistoriaBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbNroHistoriaBusqueda
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbNroHistoriaBusqueda_LostFocus()
    'cmbNroHistoriaBusqueda.Text = mo_Teclado.CapitalizarNombres(cmbNroHistoriaBusqueda.Text)
    
    If cmbNroHistoriaBusqueda.Text <> "" Then
        If Not IsNumeric(cmbNroHistoriaBusqueda.Text) Then
            MsgBox "Ha ingresado un valor no válido en la historia clínica", vbInformation, Me.Caption
            cmbNroHistoriaBusqueda.Text = ""
        End If
    End If
    
    'mo_Formulario.MarcarComoVacio cmbNroHistoriaBusqueda
End Sub

Private Sub cmbNroHistoriaBusqueda_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
        If Len(cmbNroHistoriaBusqueda.Text) >= 10 Then
            KeyAscii = 0
        End If
    Else
        If KeyAscii = vbKeyReturn Then
            cmbNroHistoriaBusqueda_Click
        End If
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
End Sub

Private Sub cmbCondicionAlta_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub cmbIdDestinoAtencion_Click()
    
    If Val(mo_cmbIdDestinoAtencion.BoundText) <> 5 And Val(mo_cmbIdDestinoAtencion.BoundText) <> 6 Then
        mo_cmbIdTipoReferenciaDestino.BoundText = ""
        Me.txtIdEstablecimientoDestino.Tag = ""
        Me.txtIdEstablecimientoDestino = ""
        Me.txtNombreDestinoReferencia = ""
    End If
    
    Select Case Val(mo_cmbIdDestinoAtencion.BoundText)
    Case 0, 1, 2, 3, 4
        Me.btnBuscarEstablecimientoDestino.Visible = False
        Me.fraDatosReferenciaDestino.Visible = False
    Case 5
        Me.btnBuscarEstablecimientoDestino.Visible = True
        Me.fraDatosReferenciaDestino.Visible = True
        Me.fraDatosReferenciaDestino = "Datos de referencia destino "
        Me.lblIdTipoReferenciaDestino = "Tipo Referencia"
        Me.lblIdEstablecimientoDestino = "Estab. Referencia"
    Case 6
        Me.btnBuscarEstablecimientoDestino.Visible = True
        Me.fraDatosReferenciaDestino.Visible = True
        Me.fraDatosReferenciaDestino = "Datos de contrareferencia destino "
        Me.lblIdTipoReferenciaDestino = "Tipo Contrarefer."
        Me.lblIdEstablecimientoDestino = "Estab. Contrarefer."
    End Select

End Sub
Private Sub cmbIdDestinoAtencion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdDestinoAtencion
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdDestinoAtencion_LostFocus()
   If cmbIdDestinoAtencion.Text <> "" Then
       mo_cmbIdDestinoAtencion.BoundText = Val(Split(cmbIdDestinoAtencion.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdDestinoAtencion
End Sub

Private Sub cmbIdDestinoAtencion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub cmbIdFuenteFinanciamiento_Click()
        
        mo_cmbIdPlan.BoundColumn = "IdPlan"
        mo_cmbIdPlan.ListField = "DescripcionLarga"
        Set mo_cmbIdPlan.RowSource = mo_AdminFacturacion.PlanesFinanciamientoSeleccionarPorTipoYFuenteFinanciamiento(Val(mo_cmbIdTipoFinanciamiento.BoundText), Val(mo_cmbIdFuenteFinanciamiento.BoundText))

End Sub

Private Sub cmbIdFuenteFinanciamiento_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdFuenteFinanciamiento
AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdFuenteFinanciamiento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub cmbIdTipoFinanciamiento_Click()
Dim sMensaje As String

    mo_cmbIdFuenteFinanciamiento.BoundText = ""
    mo_cmbIdPlan.BoundText = ""
    Me.txtNroAutorizacion = ""
    Me.txtNroPlaca = ""
    
    If mo_cmbIdTipoFinanciamiento.BoundText <> "" Then
        If mo_cmbIdTipoFinanciamiento.BoundText <> 1 Then
            mo_cmbIdFuenteFinanciamiento.BoundColumn = "IdFuenteFinanciamiento"
            mo_cmbIdFuenteFinanciamiento.ListField = "DescripcionLarga"
            Set mo_cmbIdFuenteFinanciamiento.RowSource = mo_AdminFacturacion.FuentesFinanciamientoSeleccionarPorTipo(Val(mo_cmbIdTipoFinanciamiento.BoundText))
            sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError
        
            mo_cmbIdPlan.BoundColumn = "IdPlan"
            mo_cmbIdPlan.ListField = "DescripcionLarga"
            Set mo_cmbIdPlan.RowSource = mo_AdminFacturacion.PlanesFinanciamientoSeleccionarPorTipoFinanciamiento(Val(mo_cmbIdTipoFinanciamiento.BoundText))
            sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError
          
            mo_Formulario.HabilitarDeshabilitar Me.cmbIdFuenteFinanciamiento, True
            mo_Formulario.HabilitarDeshabilitar Me.cmbIdPlan, True
            mo_Formulario.HabilitarDeshabilitar Me.txtNroAutorizacion, True
          
          Else
            
            mo_Formulario.HabilitarDeshabilitar Me.cmbIdFuenteFinanciamiento, False
            mo_Formulario.HabilitarDeshabilitar Me.cmbIdPlan, False
            mo_Formulario.HabilitarDeshabilitar Me.txtNroAutorizacion, False
          End If
    End If
    
    
    mo_Formulario.VisibleNoVisible Me.fraDatosSeguro, mo_cmbIdTipoFinanciamiento.BoundText <> 1
    mo_Formulario.VisibleNoVisible Me.txtNroPlaca, mo_cmbIdTipoFinanciamiento.BoundText = 3
    mo_Formulario.VisibleNoVisible Me.lblNroPlaca, mo_cmbIdTipoFinanciamiento.BoundText = 3


End Sub

Private Sub cmbIdTipoFinanciamiento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoFinanciamiento
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdTipoFinanciamiento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdTipoFinanciamiento_LostFocus()
   If cmbIdTipoFinanciamiento.Text <> "" Then
       mo_cmbIdTipoFinanciamiento.BoundText = Val(Split(cmbIdTipoFinanciamiento.Text, " = ")(0))
   End If
End Sub

Private Sub cmbIdTipoGravedad_Change()
    If cmbIdTipoGravedad.Text <> "" Then
        mo_cmbIdTipoGravedad.BoundText = Val(Split(cmbIdTipoGravedad.Text, " = ")(0))
        If mo_cmbIdTipoGravedad.BoundText = 5 Then
            Me.ucPacientesDetalle1.PacienteNoIdentificado = True
        Else
            Me.ucPacientesDetalle1.PacienteNoIdentificado = False
       End If
    End If
End Sub

Private Sub cmbIdTipoGravedad_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoGravedad
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdTipoGravedad_LostFocus()
   If cmbIdTipoGravedad.Text <> "" Then
       mo_cmbIdTipoGravedad.BoundText = Val(Split(cmbIdTipoGravedad.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdTipoGravedad
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


Private Sub cmbIdViasAdmision_Click()
    
    If Val(mo_cmbIdViasAdmision.BoundText) <> 5 And Val(mo_cmbIdViasAdmision.BoundText) <> 6 Then
        mo_cmbIdTipoReferenciaOrigen.BoundText = ""
        Me.txtIdEstablecimientoOrigen = ""
        Me.txtIdEstablecimientoOrigen.Tag = ""
        Me.txtNombreOrigenReferencia = ""
    End If
    
    Select Case Val(mo_cmbIdViasAdmision.BoundText)
    Case 0, 1, 2, 3, 4
        Me.btnBuscarEstablecimiento.Visible = False
        Me.fraDatosReferenciaOrigen.Visible = False
    Case 5
        Me.btnBuscarEstablecimiento.Visible = True
        Me.fraDatosReferenciaOrigen.Visible = True
        Me.fraDatosReferenciaOrigen = "Datos de referencia origen "
        Me.lblIdTipoReferenciaOrigen = "Tipo Referencia"
        Me.lblIdEstablecimientoOrigen = "Estab. Referencia"
    Case 6
        Me.btnBuscarEstablecimiento.Visible = True
        Me.fraDatosReferenciaOrigen.Visible = True
        Me.fraDatosReferenciaOrigen = "Datos de contrareferencia origen "
        Me.lblIdTipoReferenciaOrigen = "Tipo Contrarefer."
        Me.lblIdEstablecimientoOrigen = "Estab. Contrarefer."
    End Select

End Sub
Private Sub cmbIdViasAdmision_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdViasAdmision
AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdViasAdmision_LostFocus()
   If cmbIdViasAdmision.Text <> "" Then
       mo_cmbIdViasAdmision.BoundText = Val(Split(cmbIdViasAdmision.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdViasAdmision
End Sub

Private Sub cmbIdViasAdmision_KeyPress(KeyAscii As Integer)
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

Private Sub Label16_Click()

End Sub

Private Sub Form_Initialize()
    
    Set mo_cmbIdTipoServicio.MiComboBox = cmbIdTipoServicio
    Set mo_cmbIdViasAdmision.MiComboBox = cmbIdViasAdmision
    Set mo_cmbIdDestinoAtencion.MiComboBox = cmbIdDestinoAtencion
    Set mo_cmbIdTipoFinanciamiento.MiComboBox = cmbIdTipoFinanciamiento
    Set mo_cmbIdCondicionEnElServicio.MiComboBox = cmbIdCondicionEnElServicio
    Set mo_cmbIdTipoReferenciaOrigen.MiComboBox = cmbIdTipoReferenciaOrigen
    Set mo_cmbIdTipoReferenciaDestino.MiComboBox = cmbIdTipoReferenciaDestino
    Set mo_cmbIdFuenteFinanciamiento.MiComboBox = cmbIdFuenteFinanciamiento
    Set mo_cmbIdPlan.MiComboBox = cmbIdPlan
    Set mo_cmbIdCondicionEnElEstablecimiento.MiComboBox = cmbIdCondicionEnElEstablecimiento
    Set mo_cmbIdTipoGravedad.MiComboBox = cmbIdTipoGravedad
    Set mo_cmbCondicionAlta.MiComboBox = cmbCondicionAlta
    Set mo_cmbTipoAlta.MiComboBox = cmbTipoAlta

End Sub

Private Sub txtFechaEgreso_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFechaEgreso
AdministrarKeyPreview KeyCode
End Sub

Private Sub txtFechaEgreso_LostFocus()
       
       If txtFechaEgreso <> SIGHComun.FECHA_VACIA_DMY Then
            If Not EsFecha(txtFechaEgreso, "DD/MM/AAAA") Then
                MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
                 txtFechaEgreso = SIGHComun.FECHA_VACIA_DMY
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

Private Sub txtHoraEgreso_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtHoraEgreso
AdministrarKeyPreview KeyCode
End Sub

Private Sub txtHoraEgreso_LostFocus()
       
        If txtHoraEgreso <> SIGHComun.HORA_VACIA_HM Then
            If Not SIGHComun.ValidaHora(txtHoraEgreso) Then
                MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
                 txtHoraEgreso = SIGHComun.FECHA_VACIA_DMY
            End If
        End If
        
    mo_Formulario.MarcarComoVacio txtHoraEgreso
End Sub

Private Sub txtHoraEgreso_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaFecha(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtIdEstablecimientoDestino_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtIdEstablecimientoDestino
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdEstablecimientoDestino_LostFocus()
   CompletarDatosDelEstablecimientoEnElLostFocus txtIdEstablecimientoDestino, txtNombreDestinoReferencia, Val(mo_cmbIdTipoReferenciaDestino.BoundText)
   mo_Formulario.MarcarComoVacio txtIdEstablecimientoDestino
End Sub

Private Sub txtIdEstablecimientoDestino_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtApellidoMaternoBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtApellidoMaternoBusqueda
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtApellidoMaternoBusqueda_LostFocus()
txtApellidoMaternoBusqueda.Text = mo_Teclado.CapitalizarNombres(txtApellidoMaternoBusqueda.Text)
   'mo_Formulario.MarcarComoVacio txtApellidoMaternoBusqueda
End Sub

Private Sub txtApellidoMaternoBusqueda_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtFechaIngreso_Change()
    
    On Error Resume Next
    Me.txtEdadEnDias = ""
    Me.txtEdadEnDias = CalcularEdad(CDate(Me.ucPacientesDetalle1.FechaNacimiento), CDate(txtFechaIngreso))
        
End Sub

Function CalcularEdad(daFechaNacimiento As Date, daFechaReferencia As Date) As Integer
Dim iEdad As Integer

    iEdad = DateDiff("yyyy", daFechaNacimiento, daFechaReferencia)
    
    If CDate((day(daFechaNacimiento) & "/" & Month(daFechaNacimiento) & "/" & Year(daFechaReferencia))) > daFechaReferencia Then
        iEdad = iEdad - 1
    End If
    
    CalcularEdad = iEdad
    
End Function


Private Sub txtFechaIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFechaIngreso
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtFechaIngreso_LostFocus()
       
       If txtFechaIngreso <> SIGHComun.FECHA_VACIA_DMY Then
            If Not EsFecha(txtFechaIngreso, "DD/MM/AAAA") Then
                MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
                 txtFechaIngreso = SIGHComun.FECHA_VACIA_DMY
            End If
        End If
        
    Me.ucProcedimientoDetalle1.FechaIngreso = CDate(Me.txtFechaIngreso.Text + " " + Me.txtHoraIngreso.Text)
    Me.ucExamenDetalle1.FechaIngreso = CDate(Me.txtFechaIngreso.Text + " " + Me.txtHoraIngreso.Text)

       mo_Formulario.MarcarComoVacio txtFechaIngreso
End Sub

Private Sub txtFechaIngreso_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaFecha(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtHoraIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtHoraIngreso
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtHoraIngreso_LostFocus()

        If txtHoraIngreso <> SIGHComun.HORA_VACIA_HM Then
            If Not SIGHComun.ValidaHora(txtHoraIngreso) Then
                MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
                 txtHoraIngreso = SIGHComun.FECHA_VACIA_DMY
            End If
        End If
        
    Me.ucProcedimientoDetalle1.FechaIngreso = CDate(Me.txtFechaIngreso.Text + " " + Me.txtHoraIngreso.Text)
    Me.ucExamenDetalle1.FechaIngreso = CDate(Me.txtFechaIngreso.Text + " " + Me.txtHoraIngreso.Text)
        
        mo_Formulario.MarcarComoVacio txtHoraIngreso
        
End Sub

Private Sub txtHoraIngreso_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaHora(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtIdEstablecimientoOrigen_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtIdEstablecimientoOrigen
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

Private Sub txtIdMedicoEgreso_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtIdMedicoEgreso
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

Private Sub txtIdMedicoIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtIdMedicoIngreso
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdMedicoIngreso_LostFocus()
    CompletarDatosDeMedicoEnElLostFocus txtIdMedicoIngreso, lblNombreMedico
    mo_Formulario.MarcarComoVacio txtIdMedicoIngreso
End Sub

Private Sub txtIdMedicoIngreso_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtIdServicioIngreso_LostFocus()

    Me.txtIdServicioIngreso.Text = UCase(Me.txtIdServicioIngreso.Text)
    CompletarDatosDeServicioEnElLostFocus txtIdServicioIngreso, lblNombreServicio
    mo_Formulario.MarcarComoVacio txtIdServicioIngreso
    
End Sub

Private Sub txtIdServicioIngreso_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtIdServicioIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdServicioIngreso
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNroDNIBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNroDNIBusqueda
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtNroDNIBusqueda_LostFocus()
txtNroDNIBusqueda.Text = mo_Teclado.CapitalizarNombres(txtNroDNIBusqueda.Text)
   'mo_Formulario.MarcarComoVacio txtNroDNIBusqueda
End Sub

Private Sub txtNroDNIBusqueda_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtApellidoPaternoBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtApellidoPaternoBusqueda
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtApellidoPaternoBusqueda_LostFocus()
txtApellidoPaternoBusqueda.Text = mo_Teclado.CapitalizarNombres(txtApellidoPaternoBusqueda.Text)
   'mo_Formulario.MarcarComoVacio txtApellidoPaternoBusqueda
End Sub

Private Sub txtApellidoPaternoBusqueda_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtNroPlaca_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroPlaca
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtNroPlaca_LostFocus()
    txtNroPlaca = UCase(txtNroPlaca)
   mo_Formulario.MarcarComoVacio txtNroPlaca
End Sub

Private Sub txtNroPlaca_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtNroAutorizacion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNroAutorizacion
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtNroAutorizacion_LostFocus()
   mo_Formulario.MarcarComoVacio txtNroAutorizacion
End Sub

Private Sub txtNroAutorizacion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtFechaCierre_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFechaCierre
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtFechaCierre_LostFocus()
   mo_Formulario.MarcarComoVacio txtFechaCierre
End Sub

Private Sub txtFechaCierre_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaFecha(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtFechaApertura_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFechaApertura
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtFechaApertura_LostFocus()
   mo_Formulario.MarcarComoVacio txtFechaApertura
End Sub

Private Sub txtFechaApertura_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaFecha(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub cmbIdPlan_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdPlan
    AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdPlan_LostFocus()
   If cmbIdPlan.Text <> "" Then
       mo_cmbIdPlan.BoundText = Val(Split(cmbIdPlan.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdPlan
End Sub

Private Sub cmbIdPlan_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla CuentasAtencion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()

    mo_Formulario.HabilitarDeshabilitar Me.cmbIdTipoServicio, False
    
    mo_Formulario.HabilitarDeshabilitar txtFechaApertura, False
    mo_Formulario.HabilitarDeshabilitar txtHoraApertura, False
    
    mo_Formulario.HabilitarDeshabilitar Me.txtFechaCierre, False
    mo_Formulario.HabilitarDeshabilitar Me.txtHoraCierre, False
    
    mo_Formulario.VisibleNoVisible Me.fraDatosReferenciaOrigen, False
    mo_Formulario.HabilitarDeshabilitar Me.txtNombreOrigenReferencia, False
    
    mo_Formulario.VisibleNoVisible Me.fraDatosReferenciaDestino, False
    mo_Formulario.HabilitarDeshabilitar Me.txtNombreDestinoReferencia, False
    
    mo_Formulario.HabilitarDeshabilitar txtFechaApertura, False
    mo_Formulario.HabilitarDeshabilitar txtHoraApertura, False
    
    mo_Formulario.HabilitarDeshabilitar Me.txtFechaCierre, False
    mo_Formulario.HabilitarDeshabilitar Me.txtHoraCierre, False
    
    mo_Formulario.HabilitarDeshabilitar Me.lblNombreMedico, False
    mo_Formulario.HabilitarDeshabilitar Me.lblNombreServicio, False
    
    mo_Formulario.HabilitarDeshabilitar Me.lblNombreMedicoEgreso, False
    
    mo_Formulario.HabilitarDeshabilitar Me.lblCuentaAtencion, False
    
    Me.tabAdmision.TabVisible(2) = (ml_TipoVistaForm = sghVistaAtencion)
    
    Me.ucProcedimientoDetalle1.TipoServicio = sghEmergenciaConsultorios
    Me.ucExamenDetalle1.TipoServicio = sghEmergenciaConsultorios
    Me.ucPacientesDetalle1.NotaSobreUbicacion = "(*) Datos del día de la atención del paciente"
    
    Select Case mi_Opcion
        Case sghAgregar
            Me.ucPacientesDetalle1.TipoServicio = sghEmergenciaConsultorios
            ValoresPorDefecto
        Case sghModificar
            CargarDatosALosControles
        Case sghConsultar
            CargarDatosALosControles
        Case sghEliminar
            CargarDatosALosControles
    End Select

     Select Case mi_Opcion
     Case sghAgregar
        Me.btnImprimir.Enabled = False
     Case sghModificar
        fraBusqueda.Enabled = False
        Me.chkPacienteNuevo.Enabled = False
        Me.btnImprimir.Enabled = True
        If Not Me.ucPacientesDetalle1.PacienteNoIdentificado Then
            Me.ucPacientesDetalle1.HacerVisibleCheckPacienteNoIdentificado False
        End If
     Case sghConsultar
        DeshabilitarControlesParaEdicion
        Me.btnAceptar.Visible = False
        Me.btnImprimir.Enabled = True
        If Not Me.ucPacientesDetalle1.PacienteNoIdentificado Then
            Me.ucPacientesDetalle1.HacerVisibleCheckPacienteNoIdentificado False
        End If
    Case sghEliminar
        DeshabilitarControlesParaEdicion
        Me.btnImprimir.Enabled = True
        If Not Me.ucPacientesDetalle1.PacienteNoIdentificado Then
            Me.ucPacientesDetalle1.HacerVisibleCheckPacienteNoIdentificado False
        End If
    
    End Select

End Sub
Sub DeshabilitarControlesParaEdicion()
    
    fraBusqueda.Enabled = False
    fraDatosSeguro.Enabled = False
    fraDatosCuenta.Enabled = False
    Me.chkPacienteNuevo.Enabled = False
    Me.ucPacientesDetalle1.DeshabilitarFrames
    
End Sub

Sub ValoresPorDefecto()

    Me.txtFechaApertura.Text = Format(Now, "dd/mm/yyyy")
    Me.txtHoraApertura.Text = Format(Now, "hh:mm")

    Me.txtFechaIngreso.Text = Format(Now, "dd/mm/yyyy")
    Me.txtHoraIngreso = Format(Now, "hh:mm")

    mo_cmbIdTipoFinanciamiento.BoundText = "1"
    
    Me.ucPacientesDetalle1.ConfigurarValoresPorDefecto
    Me.ucProcedimientoDetalle1.FechaIngreso = CDate(Me.txtFechaIngreso.Text + " " + Me.txtHoraIngreso.Text)
    Me.ucExamenDetalle1.FechaIngreso = CDate(Me.txtFechaIngreso.Text + " " + Me.txtHoraIngreso.Text)
    

End Sub
'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla CuentasAtencion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       
    Select Case mi_Opcion
    Case sghAgregar
        Me.Caption = "Agregar Admisión Emergencia"
    Case sghModificar
        Me.Caption = "Modificar Admisión Emergencia"
    Case sghConsultar
        Me.Caption = "Consultar Admisión Emergencia"
    Case sghEliminar
        Me.Caption = "Eliminar Admisión Emergencia"
    End Select

    CargarComboBoxes
    CargarDatosAlFormulario
    mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
       
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla CuentasAtencion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Activate()
   If mi_Opcion <> sghAgregar Then
       If Not mb_ExistenDatos Then
           Me.Visible = False
       End If
   End If
End Sub
Sub AdministrarKeyPreview(KeyCode As Integer)
    
    Select Case KeyCode
    Case vbKeyEscape
        btnCancelar_Click
    Case vbKeyF2
        btnAceptar_Click
    Case vbKeyF6
            btnBuscarPaciente_Click
     Case vbKeyF7
         Me.tabAdmision.tab = 0
         Me.ucPacientesDetalle1.SetPestaniaTabPaciente 0
         On Error Resume Next
         Me.ucPacientesDetalle1.SetFocusOnDepartamentoDomicilio
     Case vbKeyF8
         Me.tabAdmision.tab = 0
         Me.ucPacientesDetalle1.SetPestaniaTabPaciente 1
         On Error Resume Next
         Me.ucPacientesDetalle1.SetFocusOnDepartamentoProcedencia
     Case vbKeyF9
         Me.tabAdmision.tab = 0
         On Error Resume Next
         Me.ucPacientesDetalle1.SetPestaniaTabPaciente 2
         Me.ucPacientesDetalle1.SetFocusOnDepartamentoNacimiento
     Case vbKeyF10
         Me.tabAdmision.tab = 0
         On Error Resume Next
         Me.ucPacientesDetalle1.SetFocusOnApellidoPaterno
     Case vbKeyF11
         Me.tabAdmision.tab = 1
         On Error Resume Next
         txtFechaIngreso.SetFocus
     Case vbKeyF12
         Me.tabAdmision.tab = 2
    End Select
       
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   'AdministrarKeyPreview KeyCode
End Sub

Private Sub btnAceptar_Click()
   Select Case mi_Opcion
   Case sghAgregar
        If ValidarDatosObligatorios() Then
            CargaDatosAlObjetosDeDatos
            If ValidarReglas() Then
               If AgregarDatos() Then
                    MsgBox "Los datos se agregaron correctamente", vbInformation, Me.Caption
                    Me.Visible = False
               Else
                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminAdmision.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
            CargaDatosAlObjetosDeDatos
           If ValidarReglas() Then
               If ModificarDatos() Then
                   MsgBox " Los datos se modificaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminAdmision.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
                CargaDatosAlObjetosDeDatos
               If EliminarDatos() Then
                   MsgBox " Los datos se eliminaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_AdminAdmision.MensajeError, vbExclamation, Me.Caption
               End If
           End If
   End Select
End Sub

Private Sub btnCancelar_Click()
   Me.Visible = False
End Sub

Function ValidarDatosObligatorios() As Boolean
   Dim sMensaje As String
   ValidarDatosObligatorios = False
   
   '-------------------------------------------------------------------------
   '                VALIDA DATOS DE LA CUENTA DE ATENCION
   '-------------------------------------------------------------------------
   'If IdCuentaAtencion = 0 Then
   '    sMensaje = sMensaje + "Ingrese el valor de IdCuentaAtencion" + Chr(13)
   'End If
   'If IdAtencion = 0 Then
   '    sMensaje = sMensaje + "Ingrese el valor de IdAtencion" + Chr(13)
   'End If
   
   If mo_cmbIdTipoFinanciamiento.BoundText = 0 Then
       sMensaje = sMensaje + "Ingrese el tipo de financiamiento" + Chr(13)
    Else
        
        If mo_cmbIdTipoFinanciamiento.BoundText <> 1 Then
            'If Val(Me.cmbIdFuenteFinanciamiento.BoundText) = 0 Then
            '    sMensaje = sMensaje + "Ingrese el valor de la fuente de financiamiento" + Chr(13)
            'End If
            If Me.txtNroAutorizacion.Text = "" Then
                sMensaje = sMensaje + "Ingrese el nro de autorización" + Chr(13)
            End If
        End If
        
        If mo_cmbIdTipoFinanciamiento.BoundText = 3 Then
            If Me.txtNroPlaca.Text = "" Then
                sMensaje = sMensaje + "Ingrese el nro de placa" + Chr(13)
            End If
        End If

   End If
   
      'If Me.txtFechaCierre.Text = SIGHcomun.FECHA_VACIA_DMY Then
   '    sMensaje = sMensaje + "Ingrese el valor de FechaCierre" + Chr(13)
   'End If
   'If Me.txtFechaApertura.Text = SIGHComun.FECHA_VACIA_DMY Then
   '    sMensaje = sMensaje + "Ingrese la fecha de apertura" + Chr(13)
   'End If
   
   'If Val(Me.cmbIdPlan.BoundText) = 0 Then
   '    sMensaje = sMensaje + "Ingrese el valor de IdPlan" + Chr(13)
   'End If
   
    '---------------------------------------------------------------------------------
    '           VALIDA DATOS DE LA ATENCION
    '---------------------------------------------------------------------------------
   ' If IdAtencion = 0 Then
   '    sMensaje = sMensaje + "Ingrese el valor de IdAtencion" + Chr(13)
   'End If
   If Me.txtIdMedicoIngreso.Text = "" Then
       sMensaje = sMensaje + "Ingrese el médico responsable de ingreso" + Chr(13)
   End If
   If Me.txtIdServicioIngreso = "" Then
       sMensaje = sMensaje + "Ingrese el servicio de ingreso" + Chr(13)
   End If
    'If Me.cmbIdTipoReferenciaOrigen.BoundText = "" Then
    '   sMensaje = sMensaje + "Ingrese el valor de IdTipoReferenciaOrigen" + Chr(13)
    'Else
    '    If Me.txtIdEstablecimientoOrigen.Text = 0 Then
    '        sMensaje = sMensaje + "Ingrese el valor de IdEstablecimientoOrigen" + Chr(13)
    '    End If
   'End If
   If Me.txtHoraIngreso.Text = "" Then
       sMensaje = sMensaje + "Ingrese la hora de ingreso" + Chr(13)
   End If
   If Me.txtFechaIngreso.Text = SIGHComun.FECHA_VACIA_DMY Then
       sMensaje = sMensaje + "Ingrese la fecha de ingreso" + Chr(13)
   End If
   If Val(mo_cmbIdTipoServicio.BoundText) = 0 Then
       sMensaje = sMensaje + "Ingrese el tipo de servicio" + Chr(13)
   End If
   If Val(Me.txtEdadEnDias.Text) = 0 Then
       sMensaje = sMensaje + "Ingrese la edad" + Chr(13)
   End If
   
    sMensaje = sMensaje + ucPacientesDetalle1.ValidarDatosObligatorios
   
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean

   ValidarReglas = False
   
    If Not Me.ucPacientesDetalle1.ValidarReglas(mo_Pacientes) Then
        Exit Function
    End If
   
'    If Me.txtFechaIngreso > Date Then
'        MsgBox "La fecha de ingreso no puede ser mayor que la fecha de hoy", vbExclamation, Me.Caption
'        Exit Function
'    End If
   
    If Me.ucPacientesDetalle1.FechaNacimiento <> SIGHComun.FECHA_VACIA_DMY Then
        If CDate(Me.ucPacientesDetalle1.FechaNacimiento) > CDate(Me.txtFechaIngreso) Then
            MsgBox "La fecha de ingreso no puede ser menor que la fecha de nacimiento", vbExclamation, Me.Caption
            Exit Function
        End If
    End If
    
    If Me.txtFechaEgreso <> SIGHComun.FECHA_VACIA_DMY Then
    
         If CDate(Me.txtFechaEgreso + " " + Me.txtHoraEgreso) < CDate(Me.txtFechaIngreso + " " + Me.txtHoraIngreso) Then
             MsgBox "La fecha de egreso no puede ser menor que la fecha de ingreso", vbExclamation, Me.Caption
             Exit Function
         End If
        
         If DateDiff("d", Me.txtFechaIngreso, Me.txtFechaEgreso) > CLng(SIGHComun.EstanciaMaxHospitalizacion) Then
             If MsgBox("¡El intervalo entre la fecha de ingreso y egreso es de mayor que la estancia máxima (" + SIGHComun.EstanciaMaxHospitalizacion + " días) ¡" + Chr(13) + "Fecha de Ingreso: " + Me.txtFechaIngreso + " -  Fecha de Egreso: " + Me.txtFechaEgreso + Chr(13) + "¿Es correcto?", vbYesNo + vbQuestion, Me.Caption) = vbNo Then
                Exit Function
            End If
         End If
    
    End If
   
   
   ValidarReglas = True

End Function

'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla CuentasAtencion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()

    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LA CUENTA ATENCION
    '---------------------------------------------------------------------------------
   With mo_CuentasAtencion
           .IdCuentaAtencion = Me.IdCuentaAtencion
           .IdAtencion = Me.IdAtencion
           .IdFuenteFinanciamiento = Val(mo_cmbIdFuenteFinanciamiento.BoundText)
            Select Case mi_Opcion
            Case sghAgregar
                .IdEstado = 1   'Cuenta Abierta
                fraDatosCuenta = fraDatosCuenta + " - ATENCIÓN ABIERTA"
                Me.chkAnularAtencion.Caption = "¿Anular atención?"
            Case Else
                If Me.chkAnularAtencion = 1 Then
                    .IdEstado = 3   'Cuenta cerrada
                    fraDatosCuenta = fraDatosCuenta + " - ATENCIÓN ANULADA"
                Else
                    .IdEstado = ml_EstadoCuenta
                End If
            End Select
            
           .IdTipoFinanciamiento = mo_cmbIdTipoFinanciamiento.BoundText
           .NroPlaca = Me.txtNroPlaca.Text
           .NroAutorizacion = Me.txtNroAutorizacion.Text
           .FechaCierre = IIf(Me.txtFechaCierre.Text = SIGHComun.FECHA_VACIA_DMY Or Me.txtFechaCierre.Text = "", 0, Me.txtFechaCierre.Text)
           .FechaApertura = IIf(Me.txtFechaApertura.Text = SIGHComun.FECHA_VACIA_DMY Or Me.txtFechaApertura.Text = "", 0, Me.txtFechaApertura.Text)
           .IdPaciente = Me.IdPaciente
           .IdPlan = Val(mo_cmbIdPlan.BoundText)
           .HoraApertura = Me.txtHoraApertura
           .HoraCierre = Me.txtHoraCierre
           .IdUsuarioAuditoria = ml_IdUsuario
   End With
   
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LA ATENCION
    '---------------------------------------------------------------------------------
   With mo_Atenciones
           .IdAtencion = Me.IdAtencion
           .IdEspecialidadMedico = ml_IdEspecialidad
           .IdMedicoIngreso = Val(Me.txtIdMedicoIngreso.Tag)
           .IdMedicoEgreso = Val(Me.txtIdMedicoEgreso.Tag)
           
           .IdServicioIngreso = Val(Me.txtIdServicioIngreso.Tag)
           .IdViaAdmision = Val(mo_cmbIdViasAdmision.BoundText)
           .IdTipoReferenciaOrigen = Val(mo_cmbIdTipoReferenciaOrigen.BoundText)
            If .IdTipoReferenciaOrigen = 1 Then
                .IdEstablecimientoOrigen = Val(Me.txtIdEstablecimientoOrigen.Tag)
                .IdEstablecimientoNoMinsaOrigen = 0
            Else
                .IdEstablecimientoOrigen = 0
                .IdEstablecimientoNoMinsaOrigen = Val(Me.txtIdEstablecimientoOrigen.Tag)
            End If
           
           .IdDestinoAtencion = Val(mo_cmbIdDestinoAtencion.BoundText)
           .IdTipoReferenciaDestino = Val(mo_cmbIdTipoReferenciaDestino.BoundText)
            If .IdTipoReferenciaDestino = 1 Then
                .IdEstablecimientoDestino = Val(Me.txtIdEstablecimientoDestino.Tag)
                .IdEstablecimientoNoMinsaDestino = 0
            Else
                .IdEstablecimientoDestino = 0
                .IdEstablecimientoNoMinsaDestino = Val(Me.txtIdEstablecimientoDestino.Tag)
            End If
           
           
           .HoraIngreso = IIf(Me.txtHoraIngreso.Text = SIGHComun.HORA_VACIA_HM, "", Me.txtHoraIngreso.Text)
           .FechaIngreso = IIf(Me.txtFechaIngreso.Text = SIGHComun.HORA_VACIA_HM, "", Me.txtFechaIngreso.Text)
           .FechaEgreso = IIf(Me.txtFechaEgreso = SIGHComun.FECHA_VACIA_DMY, 0, Me.txtFechaEgreso)
           .HoraEgreso = IIf(Me.txtHoraEgreso = SIGHComun.HORA_VACIA_HM, "", Me.txtHoraEgreso)
           .IdTipoServicio = mo_cmbIdTipoServicio.BoundText
           .Edad = Me.txtEdadEnDias.Text
           .IdPaciente = Me.IdPaciente
           .IdUsuarioAuditoria = Me.IdUsuario
           .Observacion = Me.txtObservacion
               
           'Estos datos llenaran  en el modulo de registro de atenciones
            If Me.chkPacienteNuevo = 1 Then
                .IdTipoCondicionALEstab = 1
                .IdTipoCondicionAlServicio = 1
            Else
                .IdTipoCondicionALEstab = mo_AdminServiciosComunes.TiposCondicionPacienteCondicionAlEstablecimiento(Me.IdPaciente, Format(Me.txtFechaIngreso, "dd/mm/yyyy"), Me.IdAtencion)
                .IdTipoCondicionAlServicio = mo_AdminServiciosComunes.TiposCondicionPacienteCondicionAlServicio(Me.IdPaciente, Format(Me.txtFechaIngreso, "dd/mm/yyyy"), Me.txtIdServicioIngreso.Tag, Me.IdAtencion)
            End If
            
            .FechaEgresoAdministrativo = 0
            .HoraEgresoAdministrativo = ""
            .IdCamaIngreso = 0
            .IdCamaEgreso = 0
            .IdCondicionAlta = Val(mo_cmbCondicionAlta.BoundText)
            .IdServicioEgreso = 0
            .IdTipoAlta = Val(mo_cmbTipoAlta.BoundText)
            .TieneNecropsia = 0
            .HuboInfeccionIntraHospitalaria = 0
            .IdTipoGravedad = Val(mo_cmbIdTipoGravedad.BoundText)
            .IdUsuarioAuditoria = ml_IdUsuario
   End With

    '---------------------------------------------------------------------------------
    '           CARGA DATOS PACIENTES
    '---------------------------------------------------------------------------------
    Me.ucPacientesDetalle1.IdUsuario = ml_IdUsuario
    Me.ucPacientesDetalle1.CargarDatosAlObjetoDatos mo_Pacientes, mo_Historia
    '---------------------------------------------------------------------------------
    '           COMPLETA LOS DATOS DE LA ATENCION
    '---------------------------------------------------------------------------------
   With mo_Atenciones
            'Datos de ubicación de paciente en esta atencion
            .IdPaisDomicilio = mo_Pacientes.IdPaisDomicilio
            .IdDepartamentoDomicilio = mo_Pacientes.IdDepartamentoDomicilio
            .IdProvinciaDomicilio = mo_Pacientes.IdProvinciaDomicilio
            .IdDistritoDomicilio = mo_Pacientes.IdDistritoDomicilio
            .IdCentroPobladoDomicilio = mo_Pacientes.IdCentroPobladoDomicilio
            
            .IdPaisProcedencia = mo_Pacientes.IdPaisProcedencia
            .IdDepartamentoProcedencia = mo_Pacientes.IdDepartamentoProcedencia
            .IdProvinciaProcedencia = mo_Pacientes.IdProvinciaProcedencia
            .IdDistritoProcedencia = mo_Pacientes.IdDistritoProcedencia
            .IdCentroPobladoProcedencia = mo_Pacientes.IdCentroPobladoProcedencia
            
            .DireccionDomicilio = mo_Pacientes.DireccionDomicilio
            .EtapaDomicilio = mo_Pacientes.EtapaDomicilio
            .SectorDomicilio = mo_Pacientes.SectorDomicilio
            .LoteDomicilio = mo_Pacientes.LoteDomicilio
            .ManzanaDomicilio = mo_Pacientes.ManzanaDomicilio
            .PisoDomicilio = mo_Pacientes.PisoDomicilio
            .NroDomicilio = mo_Pacientes.NroDomicilio
    End With


    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE DIAGNOSTICOS DE INGRESO
    '---------------------------------------------------------------------------------
    Me.ucDiagnosticoDetalle1.IdUsuario = ml_IdUsuario
    Me.ucDiagnosticoDetalle1.CargarDiagnosticosAlObjetoDatos mo_Diagnosticos
    
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE PROCEDIMIENTOS DE INGRESO
    '---------------------------------------------------------------------------------
    Me.ucProcedimientoDetalle1.IdUsuario = ml_IdUsuario
    Me.ucProcedimientoDetalle1.IdInterconsulta = 0
    Me.ucProcedimientoDetalle1.CargarProcedimientosAlObjetoDatos mo_Procedimientos

    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE EXAMENES DE INGRESO
    '---------------------------------------------------------------------------------
    Me.ucExamenDetalle1.IdUsuario = ml_IdUsuario
    Me.ucExamenDetalle1.CargarExamensAlObjetoDatos mo_Examenes

End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean

    AgregarDatos = mo_AdminAdmision.AdmisionEmergAgregar(mo_CuentasAtencion, mo_Atenciones, mo_Pacientes, mo_Historia, Me.ucPacientesDetalle1.TipoGeneracionHistoriaClinicaAnterior, mo_Diagnosticos, mo_Procedimientos, mo_Examenes, Me.ucPacientesDetalle1.IdHistoriaClinicaAnterior)

End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean

    ModificarDatos = mo_AdminAdmision.AdmisionEmergModificar(mo_CuentasAtencion, mo_Atenciones, mo_Pacientes, mo_Historia, Me.ucPacientesDetalle1.TipoGeneracionHistoriaClinicaAnterior, mo_Diagnosticos, mo_Procedimientos, mo_Examenes, Me.ucPacientesDetalle1.IdHistoriaClinicaAnterior)
    
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean

    EliminarDatos = mo_AdminAdmision.AdmisionEmergEliminar(mo_CuentasAtencion, mo_Atenciones)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla CuentasAtencion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()
                
        '1do:   CARGAR DATOS DE LA ATENCION
        CargarDatosDelaAtencion
        
        '2do:   CARGAR DATOS DE LA CUENTA DE ATENCION
        CargarDatosDeLaCuentaDeAtencion
       
        '3ro:   CARGAR DATOS DEL PACIENTE
        Me.ucPacientesDetalle1.IdPaciente = Me.IdPaciente
        Me.ucPacientesDetalle1.CargarDatosDePacienteALosControles
       
        '4to:   PARA VISUALIZAR LA UBICACION DEL PACIENTE AL DIA DE LA ATENCION
        Me.ucPacientesDetalle1.ReemplazarDatosDeUbicacion mo_DoUbicacionPaciente
       
        '4to:   CARGAR DATOS DE LOS DIAGNOSTICOS POR ATENCION
        Me.ucDiagnosticoDetalle1.IdAtencion = Me.IdAtencion
        Me.ucDiagnosticoDetalle1.IdInterconsulta = 0
        Me.ucDiagnosticoDetalle1.CargarDatosDeDiagnosticos
        
        '5to:   CARGAR DATOS DE LOS PROCEDIMIENTOS POR ATENCION
        Me.ucProcedimientoDetalle1.IdCuentaAtencion = Me.IdCuentaAtencion
        Me.ucProcedimientoDetalle1.IdInterconsulta = 0
        Me.ucProcedimientoDetalle1.CargarDatosDeProcedimientos
        
        '6Mo:   CARGAR DATOS DE LOS EXAMENES POR ATENCION
        Me.ucExamenDetalle1.IdCuentaAtencion = Me.IdCuentaAtencion
        Me.ucExamenDetalle1.CargarDatosDeExamens
       
End Sub
Sub CargarDatosDelaAtencion()
Dim oDoMedico As New DOMedico
Dim oDOEmpleado As New DOEmpleado
Dim oDOEspecialidades As New Collection
Dim oDOServicio As New DOServicio

        Set mo_Atenciones = mo_AdminAdmision.AtencionesSeleccionarPorId(Me.IdAtencion)
        If mo_AdminAdmision.MensajeError <> "" Then
             MsgBox "No se pudo obtener los datos" + Chr(13) + mo_AdminServiciosComunes.MensajeError, vbCritical, Me.Caption
             mb_ExistenDatos = False
             Exit Sub
        End If
        
        If Not mo_Atenciones Is Nothing Then
           With mo_Atenciones
                Me.IdAtencion = .IdAtencion
                Me.IdPaciente = .IdPaciente
                
                mo_cmbIdTipoServicio.BoundText = .IdTipoServicio
                mo_cmbIdDestinoAtencion.BoundText = .IdDestinoAtencion
                
                Me.txtIdServicioIngreso.Tag = .IdServicioIngreso
                Me.txtIdMedicoIngreso.Tag = .IdMedicoIngreso
                Me.txtIdMedicoEgreso.Tag = .IdMedicoEgreso
                
                Me.IdEspecialidad = .IdEspecialidadMedico
                Me.txtObservacion = .Observacion
                mo_cmbIdViasAdmision.BoundText = .IdViaAdmision
                mo_cmbIdTipoReferenciaOrigen.BoundText = .IdTipoReferenciaOrigen
                CompletarDatosDelEstablecimientoEnElLoad .IdEstablecimientoOrigen, .IdEstablecimientoNoMinsaOrigen, txtIdEstablecimientoOrigen, txtNombreOrigenReferencia, .IdTipoReferenciaOrigen
                
                mo_cmbIdDestinoAtencion.BoundText = .IdDestinoAtencion
                mo_cmbIdTipoReferenciaDestino.BoundText = .IdTipoReferenciaDestino
                CompletarDatosDelEstablecimientoEnElLoad .IdEstablecimientoDestino, .IdEstablecimientoNoMinsaDestino, txtIdEstablecimientoDestino, txtNombreDestinoReferencia, .IdTipoReferenciaDestino
                
                Me.txtHoraIngreso.Text = IIf(.HoraIngreso = "", SIGHComun.HORA_VACIA_HM, .HoraIngreso)
                Me.txtFechaIngreso.Text = IIf(.FechaIngreso = 0, SIGHComun.FECHA_VACIA_DMY, .FechaIngreso)
                
                Me.txtHoraEgreso.Text = IIf(.HoraEgreso = "", SIGHComun.HORA_VACIA_HM, .HoraEgreso)
                Me.txtFechaEgreso.Text = IIf(.FechaEgreso = 0, SIGHComun.FECHA_VACIA_DMY, .FechaEgreso)
                
                Me.txtEdadEnDias.Text = .Edad
                Me.txtEdadEnDias.Tag = .Edad
                
                If mo_AdminProgramacion.MedicosSeleccionarPorId(.IdMedicoIngreso, oDoMedico, oDOEmpleado, oDOEspecialidades) Then
                    Me.txtIdMedicoIngreso = oDOEmpleado.CodigoPlanilla
                    Me.lblNombreMedico = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
                Else
                    Me.lblNombreMedico = ""
                End If
                
                If mo_AdminProgramacion.MedicosSeleccionarPorId(.IdMedicoEgreso, oDoMedico, oDOEmpleado, oDOEspecialidades) Then
                    Me.txtIdMedicoEgreso = oDOEmpleado.CodigoPlanilla
                    Me.lblNombreMedicoEgreso = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
                Else
                    Me.lblNombreMedicoEgreso = ""
                End If
                
                Set oDOServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(.IdServicioIngreso)
                If Not oDOServicio Is Nothing Then
                    Me.txtIdServicioIngreso.Tag = oDOServicio.IdServicio
                    Me.txtIdServicioIngreso.Text = oDOServicio.Codigo
                    Me.lblNombreServicio = oDOServicio.Nombre
                    mo_cmbIdTipoServicio.BoundText = oDOServicio.IdTipoServicio
                    Me.IdEspecialidad = oDOServicio.IdEspecialidad
                Else
                    Me.txtIdServicioIngreso.Tag = ""
                    Me.lblNombreServicio = ""
                    mo_cmbIdTipoServicio.BoundText = ""
                End If
                
                mo_cmbIdCondicionEnElServicio.BoundText = .IdTipoCondicionAlServicio
                mo_cmbIdCondicionEnElEstablecimiento.BoundText = .IdTipoCondicionALEstab
                
                mo_cmbCondicionAlta.BoundText = .IdCondicionAlta
                mo_cmbTipoAlta.BoundText = .IdTipoAlta
                mo_cmbIdTipoGravedad.BoundText = .IdTipoGravedad
                
                If Me.txtFechaIngreso.Text <> SIGHComun.FECHA_VACIA_DMY Then
                    Me.ucProcedimientoDetalle1.FechaIngreso = CDate(Me.txtFechaIngreso.Text + " " + Me.txtHoraIngreso.Text)
                    Me.ucExamenDetalle1.FechaIngreso = CDate(Me.txtFechaIngreso.Text + " " + Me.txtHoraIngreso.Text)
                End If
                'ESTOS DATOS SE UTILIZARAN MAS ADELANTE PARA ACTUALIZAR LA UBICACION DE PACIENTE
                mo_DoUbicacionPaciente.IdPaisDomicilio = .IdPaisDomicilio
                mo_DoUbicacionPaciente.IdDepartamentoDomicilio = .IdDepartamentoDomicilio
                mo_DoUbicacionPaciente.IdProvinciaDomicilio = .IdProvinciaDomicilio
                mo_DoUbicacionPaciente.IdDistritoDomicilio = .IdDistritoDomicilio
                mo_DoUbicacionPaciente.IdCentroPobladoDomicilio = .IdCentroPobladoDomicilio
                
                mo_DoUbicacionPaciente.IdPaisProcedencia = .IdPaisProcedencia
                mo_DoUbicacionPaciente.IdDepartamentoProcedencia = .IdDepartamentoProcedencia
                mo_DoUbicacionPaciente.IdProvinciaProcedencia = .IdProvinciaProcedencia
                mo_DoUbicacionPaciente.IdDistritoProcedencia = .IdDistritoProcedencia
                mo_DoUbicacionPaciente.IdCentroPobladoProcedencia = .IdCentroPobladoProcedencia
                
                mo_DoUbicacionPaciente.DireccionDomicilio = .DireccionDomicilio
                mo_DoUbicacionPaciente.NroDomicilio = .NroDomicilio
                mo_DoUbicacionPaciente.ManzanaDomicilio = .ManzanaDomicilio
                mo_DoUbicacionPaciente.LoteDomicilio = .LoteDomicilio
                mo_DoUbicacionPaciente.SectorDomicilio = .SectorDomicilio
                mo_DoUbicacionPaciente.EtapaDomicilio = .EtapaDomicilio
                mo_DoUbicacionPaciente.PisoDomicilio = .PisoDomicilio
                
                mb_ExistenDatos = True
           End With
       Else
           mb_ExistenDatos = False
           Exit Sub
       End If

End Sub
Sub CargarDatosDeLaCuentaDeAtencion()
       
       Me.IdCuentaAtencion = mo_AdminFacturacion.CuentasAtencionSeleccionarIdPorIdAtencion(Me.IdAtencion)
       If Me.IdCuentaAtencion = 0 Then
            mb_ExistenDatos = False
            Exit Sub
       End If
        
        Me.lblCuentaAtencion = Me.IdCuentaAtencion
        
       Set mo_CuentasAtencion = mo_AdminFacturacion.CuentasAtencionSeleccionarPorId(Me.IdCuentaAtencion)
        If mo_AdminFacturacion.MensajeError <> "" Then
             MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminServiciosComunes.MensajeError, vbCritical, Me.Caption"
             mb_ExistenDatos = False
             Exit Sub
        End If
        If Not mo_CuentasAtencion Is Nothing Then
            With mo_CuentasAtencion
                Me.IdCuentaAtencion = .IdCuentaAtencion
                mo_cmbIdTipoFinanciamiento.BoundText = .IdTipoFinanciamiento    'esto debe estar antes del resto
                mo_cmbIdFuenteFinanciamiento.BoundText = .IdFuenteFinanciamiento
                Me.txtNroPlaca.Text = .NroPlaca
                Me.txtNroAutorizacion.Text = .NroAutorizacion
                Me.txtFechaCierre.Text = IIf(.FechaCierre = 0, SIGHComun.FECHA_VACIA_DMY, .FechaCierre)
                Me.txtFechaApertura.Text = IIf(.FechaApertura = 0, SIGHComun.FECHA_VACIA_DMY, .FechaApertura)
                Me.txtHoraApertura = IIf(.HoraApertura = "", SIGHComun.HORA_VACIA_HM, .HoraApertura)
                Me.txtHoraCierre = IIf(.HoraCierre = "", SIGHComun.HORA_VACIA_HM, .HoraCierre)
                mo_cmbIdPlan.BoundText = .IdPlan
                ml_EstadoCuenta = .IdEstado
                
                Select Case .IdEstado
                Case 1
                    fraDatosCuenta = fraDatosCuenta + " - ATENCIÓN ABIERTA"
                Case 2
                    fraDatosCuenta = fraDatosCuenta + " - ATENCIÓN CERRADA"
                    Me.chkAnularAtencion.Enabled = False
                    Me.btnAceptar.Enabled = False
                Case 3
                    fraDatosCuenta = fraDatosCuenta + " - ATENCIÓN ANULADA"
                    Me.chkAnularAtencion.Value = 1
                    Me.chkAnularAtencion.Enabled = False
                    Me.btnAceptar.Enabled = False
                End Select
                
                mb_ExistenDatos = True
            End With
        Else
            mb_ExistenDatos = False
            Exit Sub
        End If

End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla CuentasAtencion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()

           'LIMPIAR DATOS DE LA CUENTA DE ATENCION
           Me.IdCuentaAtencion = 0
           Me.IdAtencion = 0
           mo_cmbIdFuenteFinanciamiento.BoundText = ""
           mo_cmbIdTipoFinanciamiento.BoundText = "1"
           Me.txtNroPlaca.Text = ""
           Me.txtNroAutorizacion.Text = ""
           'Me.txtFechaCierre.Text = SIGHComun.FECHA_VACIA_DMY
           'Me.txtFechaApertura.Text = SIGHComun.FECHA_VACIA_DMY
           mo_cmbIdPlan.BoundText = ""
           'Me.txtHoraApertura = SIGHComun.HORA_VACIA_HM
           'Me.txtHoraCierre = SIGHComun.HORA_VACIA_HM
           Me.lblCuentaAtencion = ""
           
           'LIMPIAR DATOS DE LA ATENCION
           Me.IdAtencion = 0
           'Me.cmbIdEspecialidadMedico.BoundText = ""
           'Me.txtMedico.Text = ""
           'Me.cmbIdServicio.Text = ""
           'Me.cmbIdTipoCondicionALEstab.BoundText = ""
           'Me.cmbIdTipoCondicionAlServicio.BoundText = ""
           'Me.cmbIdDestinoAtencion.BoundText = ""
           'Me.cmbIdTipoReferenciaDestino.BoundText = ""
           'Me.txtIdEstablecimientoDestino.Text = ""
           'Me.txtHoraIngreso.Text = SIGHComun.HORA_VACIA_HM
           'Me.txtFechaIngreso.Text = SIGHComun.FECHA_VACIA_DMY
           'Me.cmbIdTipoServicio.BoundText = ""
           Me.txtEdadEnDias.Text = ""
           
           Me.ucPacientesDetalle1.LimpiarDatosDePaciente
           
End Sub
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'*****************************************************************************************
'                               EVENTOS DE LA ATENCION
'*****************************************************************************************
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
Private Sub cmbIdTipoServicio_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoServicio
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdTipoServicio_LostFocus()
   If cmbIdTipoServicio.Text <> "" Then
       mo_cmbIdTipoServicio.BoundText = Val(Split(cmbIdTipoServicio.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdTipoServicio
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

Private Sub txtPrimerNombreBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtPrimerNombreBusqueda
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtPrimerNombreBusqueda_LostFocus()
txtPrimerNombreBusqueda.Text = mo_Teclado.CapitalizarNombres(txtPrimerNombreBusqueda.Text)
   'mo_Formulario.MarcarComoVacio txtPrimerNombreBusqueda
End Sub

Private Sub txtPrimerNombreBusqueda_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtSegundoNombreBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtSegundoNombreBusqueda
AdministrarKeyPreview KeyCode
End Sub

Private Sub txtSegundoNombreBusqueda_LostFocus()
txtSegundoNombreBusqueda.Text = mo_Teclado.CapitalizarNombres(txtSegundoNombreBusqueda.Text)
   'mo_Formulario.MarcarComoVacio txtSegundoNombreBusqueda
End Sub

Private Sub txtSegundoNombreBusqueda_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub cmbIdFuenteFinanciamiento_LostFocus()
   
   If cmbIdFuenteFinanciamiento.Text <> "" Then
       mo_cmbIdFuenteFinanciamiento.BoundText = Val(Split(cmbIdFuenteFinanciamiento.Text, " = ")(0))
   End If

End Sub
Private Sub txtObservacion_LostFocus()
   'mo_Formulario.MarcarComoVacio txtObservacion
End Sub

Private Sub txtObservacion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub ucPacientesDetalle1_SeModificoFechaNacimiento(sFechaNacimiento As String)
    
    On Error Resume Next
    Me.txtEdadEnDias = ""
    Me.txtEdadEnDias = CalcularEdad(CDate(sFechaNacimiento), CDate(txtFechaIngreso))
    If Me.txtEdadEnDias = "" Then
        Me.txtEdadEnDias = Me.txtEdadEnDias.Tag
    End If
    
End Sub

Private Sub ucPacientesDetalle1_SeModificoPacienteNoIdentificado(bPacienteNoIdentificado As Boolean)
    If bPacienteNoIdentificado Then
        chkPacienteNuevo.Value = 1
        chkPacienteNuevo.Enabled = False
        lblObservaciones = "Datos del acompañante"
    Else
        chkPacienteNuevo.Enabled = True
        chkPacienteNuevo.Value = 1
        lblObservaciones = "Observación"
    End If
    mo_Formulario.HabilitarDeshabilitar txtObservacion, Not bPacienteNoIdentificado

End Sub

Private Sub ucPacientesDetalle1_SePresionoTeclaEspecial(KeyCode As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub CompletarDatosDeEstablecimiento(txtIdEstablecimiento As TextBox, lblNombreEstablecimiento As TextBox, lTipoReferencia As Long)
    
    If lTipoReferencia = 1 Then
        Dim oBusqueda As New EstablecimientosBusqueda
        Dim oDoEstablecimiento As New DOEstablecimiento
        oBusqueda.Show 1
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
    Else
        Dim oBusquedaNM As New EstablecimientosNoMinsaBusqueda
        Dim oDoEstablecimientoNM As New DOEstablecimientoNoMinsa
        oBusquedaNM.Show 1
        If oBusquedaNM.BotonPresionado = sghAceptar Then
            Set oDoEstablecimientoNM = mo_AdminServiciosComunes.EstablecimientosNoMinsaSeleccionarPorId(oBusquedaNM.IdRegistroSeleccionado)
            If Not oDoEstablecimientoNM Is Nothing Then
                txtIdEstablecimiento.Tag = oDoEstablecimientoNM.IdEstablecimientoNoMinsa
                txtIdEstablecimiento.Text = oDoEstablecimientoNM.IdEstablecimientoNoMinsa
                lblNombreEstablecimiento = oDoEstablecimientoNM.Nombre
            Else
                txtIdEstablecimiento.Tag = ""
                txtIdEstablecimiento.Text = ""
                lblNombreEstablecimiento = ""
            End If
        End If
    End If

End Sub
Sub CompletarDatosDelEstablecimientoEnElLostFocus(txtIdEstablecimiento As TextBox, lblNombreEstablecimiento As TextBox, lTipoReferencia As Long)
    
    If txtIdEstablecimiento <> "" Then
        If lTipoReferencia = 1 Then
                Dim oDoEstablecimiento As New DOEstablecimiento
                If mo_AdminServiciosComunes.EstablecimientosSeleccionarPorCodigo(txtIdEstablecimiento.Text, oDoEstablecimiento) Then
                    txtIdEstablecimiento.Tag = oDoEstablecimiento.IdEstablecimiento
                    txtIdEstablecimiento.Text = oDoEstablecimiento.Codigo
                    lblNombreEstablecimiento = oDoEstablecimiento.Nombre
                Else
                    txtIdEstablecimiento.Tag = ""
                    txtIdEstablecimiento = ""
                    lblNombreEstablecimiento = ""
                End If
        Else
                Dim oDOEstablecimientoNoMinsa As New DOEstablecimientoNoMinsa
                Set oDOEstablecimientoNoMinsa = mo_AdminServiciosComunes.EstablecimientosNoMinsaSeleccionarPorId(txtIdEstablecimiento.Text)
                If Not oDOEstablecimientoNoMinsa Is Nothing Then
                    txtIdEstablecimiento.Tag = oDOEstablecimientoNoMinsa.IdEstablecimientoNoMinsa
                    txtIdEstablecimiento.Text = oDOEstablecimientoNoMinsa.IdEstablecimientoNoMinsa
                    lblNombreEstablecimiento = oDOEstablecimientoNoMinsa.Nombre
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
             txtIdEstablecimiento.Text = oDOEstablecimientoNoMinsa.IdEstablecimientoNoMinsa
             txtIdEstablecimiento.Tag = oDOEstablecimientoNoMinsa.IdEstablecimientoNoMinsa
             lblNombreEstablecimiento = oDOEstablecimientoNoMinsa.Nombre
        Else
             txtIdEstablecimiento.Text = ""
             txtIdEstablecimiento.Tag = ""
             lblNombreEstablecimiento = ""
         End If
    End If

End Sub

Sub CompletarDatosDeMedico(txtMedico As TextBox, lblNombreMedico As TextBox)
Dim oBusqueda As New MedicosBusqueda
Dim oDoMedico As New DOMedico
Dim oDOEmpleado As New DOEmpleado
Dim oDOEspecialidades As New Collection

    oBusqueda.Show 1
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        If mo_AdminProgramacion.MedicosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado, oDoMedico, oDOEmpleado, oDOEspecialidades) Then
            txtMedico.Text = oDOEmpleado.CodigoPlanilla
            txtMedico.Tag = oDoMedico.IdMedico
            lblNombreMedico.Text = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        End If
    End If

End Sub
Sub CompletarDatosDeMedicoEnElLostFocus(txtMedico As TextBox, lblNombreMedico As TextBox)
Dim oMedicosEspecialidad As New Collection

    txtMedico = Trim(txtMedico)
    If txtMedico <> "" Then
        Dim oDOEmpleado As New DOEmpleado
        Dim oDoMedico As New DOMedico
        If mo_AdminProgramacion.MedicosSeleccionarPorCodigo(Str(txtMedico), oDoMedico, oDOEmpleado, oMedicosEspecialidad) Then
            txtMedico.Tag = oDoMedico.IdMedico
            Set oDOEmpleado = mo_AdminServiciosComunes.EmpleadosSeleccionarPorId(oDoMedico.IdEmpleado)
            lblNombreMedico = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        Else
            txtMedico.Tag = ""
            lblNombreMedico = ""
        End If
    End If
    
End Sub
Sub CompletarDatosDeServicio(txtIdServicio As TextBox, lblDescripcionServicio As TextBox)
Dim oBusqueda As New ServiciosBusqueda
Dim oDOServicio As New DOServicio

    oBusqueda.IdTipoServicio = Val(mo_cmbIdTipoServicio.BoundText)
    oBusqueda.HabilitarTipoServicio = False
    oBusqueda.Show 1
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
        If Not oDOServicio Is Nothing Then
            If Val(mo_cmbIdTipoServicio.BoundText) = oDOServicio.IdTipoServicio Then
                txtIdServicio.Text = oDOServicio.Codigo
                txtIdServicio.Tag = oDOServicio.IdServicio
                lblDescripcionServicio.Text = oDOServicio.Nombre
            Else
                MsgBox "El servicio seleccionado no pertenece a emergencia", vbInformation, Me.Caption
                txtIdServicio.Text = ""
                txtIdServicio.Tag = ""
                lblDescripcionServicio = ""
            End If
        End If
    End If

End Sub
Sub CompletarDatosDeServicioEnElLostFocus(txtIdServicio As TextBox, lblDescripcionServicio As TextBox)
    
    txtIdServicio.Text = UCase(txtIdServicio.Text)
    If txtIdServicio.Text <> "" Then
        Dim oDOServicio As DOServicio
        Set oDOServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorCodigo(txtIdServicio.Text)
        If Not oDOServicio Is Nothing Then
            If mo_cmbIdTipoServicio.BoundText = oDOServicio.IdTipoServicio Then
                txtIdServicio.Tag = oDOServicio.IdServicio
                lblDescripcionServicio.Text = oDOServicio.Nombre
            Else
                MsgBox "El servicio ingresado no pertenece es de emergencia", vbInformation, Me.Caption
                txtIdServicio.Tag = ""
                lblDescripcionServicio.Text = ""
            End If
        Else
            txtIdServicio.Tag = ""
            lblDescripcionServicio.Text = ""
        End If
   End If

End Sub

