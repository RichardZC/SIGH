VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.Form FacturacionEstadoCuenta 
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   11550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   90
      TabIndex        =   18
      Top             =   8340
      Width           =   11400
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "FacturacionEstadoCuenta.frx":0000
         DownPicture     =   "FacturacionEstadoCuenta.frx":0460
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
         Left            =   4305
         Picture         =   "FacturacionEstadoCuenta.frx":08D5
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "FacturacionEstadoCuenta.frx":0D4A
         DownPicture     =   "FacturacionEstadoCuenta.frx":120E
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
         Left            =   5850
         Picture         =   "FacturacionEstadoCuenta.frx":16FA
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   225
         Width           =   1365
      End
   End
   Begin UltraGrid.SSUltraGrid SSUltraGrid1 
      Height          =   6225
      Left            =   60
      TabIndex        =   17
      Top             =   2040
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   10980
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108864
      Caption         =   "SSUltraGrid1"
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
      Left            =   60
      TabIndex        =   12
      Top             =   60
      Width           =   11430
      Begin VB.TextBox txtNroCuenta 
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
         Left            =   3345
         TabIndex        =   14
         Top             =   450
         Width           =   1350
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   4770
         Picture         =   "FacturacionEstadoCuenta.frx":1BE6
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   450
         Width           =   1305
      End
      Begin PVCOMBOLibCtl.PVComboBox cmbNroHistoriaBusqueda 
         Height          =   315
         Left            =   1695
         TabIndex        =   15
         Top             =   450
         Width           =   1590
         _Version        =   524288
         _cx             =   2805
         _cy             =   556
         Appearance      =   0
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
         PrimaryColumn   =   7
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
         ColumnCount     =   12
         Column0.Heading =   "IdPaciente"
         Column0.Width   =   40
         Column0.Alignment=   0
         Column0.Hidden  =   -1  'True
         Column0.Name    =   "IdPaciente"
         Column0.Format  =   ""
         Column0.Bound   =   -1  'True
         Column0.Locked  =   0   'False
         Column0.HeaderAlignment=   0
         Column1.Heading =   "IdAtencion"
         Column1.Width   =   40
         Column1.Alignment=   0
         Column1.Hidden  =   -1  'True
         Column1.Name    =   "IdAtencion"
         Column1.Format  =   ""
         Column1.Bound   =   -1  'True
         Column1.Locked  =   0   'False
         Column1.HeaderAlignment=   0
         Column2.Heading =   "IdTipoNumeracion"
         Column2.Width   =   40
         Column2.Alignment=   0
         Column2.Hidden  =   -1  'True
         Column2.Name    =   "IdTipoNumeracion"
         Column2.Format  =   ""
         Column2.Bound   =   -1  'True
         Column2.Locked  =   0   'False
         Column2.HeaderAlignment=   0
         Column3.Heading =   "Ap. paterno"
         Column3.Width   =   50
         Column3.Alignment=   0
         Column3.Hidden  =   0   'False
         Column3.Name    =   "ApellidoPaterno"
         Column3.Format  =   ""
         Column3.Bound   =   -1  'True
         Column3.Locked  =   0   'False
         Column3.HeaderAlignment=   0
         Column4.Heading =   "Ap. materno"
         Column4.Width   =   50
         Column4.Alignment=   0
         Column4.Hidden  =   0   'False
         Column4.Name    =   "ApellidoMaterno"
         Column4.Format  =   ""
         Column4.Bound   =   -1  'True
         Column4.Locked  =   0   'False
         Column4.HeaderAlignment=   0
         Column5.Heading =   "1er Nombre"
         Column5.Width   =   50
         Column5.Alignment=   0
         Column5.Hidden  =   0   'False
         Column5.Name    =   "PrimerNombre"
         Column5.Format  =   ""
         Column5.Bound   =   -1  'True
         Column5.Locked  =   0   'False
         Column5.HeaderAlignment=   0
         Column6.Heading =   "2do Nombre"
         Column6.Width   =   50
         Column6.Alignment=   0
         Column6.Hidden  =   0   'False
         Column6.Name    =   "SegundoNombre"
         Column6.Format  =   ""
         Column6.Bound   =   -1  'True
         Column6.Locked  =   0   'False
         Column6.HeaderAlignment=   0
         Column7.Heading =   "Nro Historia Clinica"
         Column7.Width   =   40
         Column7.Alignment=   0
         Column7.Hidden  =   0   'False
         Column7.Name    =   "NroHistoriaClinica"
         Column7.Format  =   "0"
         Column7.Bound   =   -1  'True
         Column7.Locked  =   0   'False
         Column7.HeaderAlignment=   0
         Column8.Heading =   "Fecha Ingreso"
         Column8.Width   =   40
         Column8.Alignment=   0
         Column8.Hidden  =   0   'False
         Column8.Name    =   "FechaIngreso"
         Column8.Format  =   ""
         Column8.Bound   =   -1  'True
         Column8.Locked  =   0   'False
         Column8.HeaderAlignment=   0
         Column9.Heading =   "Hora Ingreso"
         Column9.Width   =   40
         Column9.Alignment=   0
         Column9.Hidden  =   0   'False
         Column9.Name    =   "HoraIngreso"
         Column9.Format  =   ""
         Column9.Bound   =   -1  'True
         Column9.Locked  =   0   'False
         Column9.HeaderAlignment=   0
         Column10.Heading=   "Nro Cuenta"
         Column10.Width  =   40
         Column10.Alignment=   0
         Column10.Hidden =   0   'False
         Column10.Name   =   "IdCuentaAtencion"
         Column10.Format =   ""
         Column10.Bound  =   -1  'True
         Column10.Locked =   0   'False
         Column10.HeaderAlignment=   0
         Column11.Heading=   "Servicio Ingreso"
         Column11.Width  =   40
         Column11.Alignment=   0
         Column11.Hidden =   0   'False
         Column11.Name   =   "ServicioIngreso"
         Column11.Format =   ""
         Column11.Bound  =   -1  'True
         Column11.Locked =   0   'False
         Column11.HeaderAlignment=   0
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
         Caption         =   "Historia clínica        Nro Cuenta"
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
         Left            =   1980
         TabIndex        =   16
         Top             =   180
         Width           =   2685
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del paciente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   60
      TabIndex        =   0
      Top             =   945
      Width           =   11430
      Begin VB.TextBox lblServicioIngreso 
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
         Left            =   7200
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   630
         Width           =   4065
      End
      Begin VB.TextBox lblFechaIngreso 
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
         Left            =   9825
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   255
         Width           =   1425
      End
      Begin VB.TextBox lblPaciente 
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
         Left            =   1695
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   645
         Width           =   4020
      End
      Begin VB.TextBox lblNroCuenta 
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
         Left            =   1695
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   255
         Width           =   1140
      End
      Begin VB.ComboBox cmbIdTipoGenHistoriaClinica 
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
         Left            =   5130
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   3315
      End
      Begin VB.TextBox txtIdNroHistoria 
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
         Left            =   3915
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label3 
         Caption         =   "Servicio Ingreso"
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
         Left            =   5805
         TabIndex        =   11
         Top             =   675
         Width           =   1305
      End
      Begin VB.Label Label7 
         Caption         =   "Nº historia"
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
         Left            =   3000
         TabIndex        =   10
         Top             =   285
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Ingreso"
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
         Left            =   8535
         TabIndex        =   9
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "Paciente"
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
         Left            =   165
         TabIndex        =   8
         Top             =   675
         Width           =   1005
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   150
         TabIndex        =   7
         Top             =   300
         Width           =   1065
      End
   End
End
Attribute VB_Name = "FacturacionEstadoCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

