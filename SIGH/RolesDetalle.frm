VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form RolesDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11835
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   11835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5955
      Left            =   60
      TabIndex        =   6
      Top             =   810
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   10504
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
      TabCaption(0)   =   "Módulos"
      TabPicture(0)   =   "RolesDetalle.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grdPermisos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Permisos"
      TabPicture(1)   =   "RolesDetalle.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "grdPermisos2"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Reportes"
      TabPicture(2)   =   "RolesDetalle.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(1)=   "grdReportes"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame5 
         Height          =   675
         Left            =   -74880
         TabIndex        =   23
         Top             =   660
         Width           =   11385
         Begin VB.CheckBox chkNingunReporte 
            Caption         =   "Ninguno"
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
            Left            =   7590
            TabIndex        =   29
            Top             =   180
            Width           =   1155
         End
         Begin VB.CheckBox chkTodosReportes 
            Caption         =   "Todos"
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
            Left            =   6600
            TabIndex        =   28
            Top             =   180
            Width           =   885
         End
         Begin VB.CommandButton cmbDelR 
            DisabledPicture =   "RolesDetalle.frx":0054
            DownPicture     =   "RolesDetalle.frx":03DF
            Height          =   315
            Left            =   10260
            Picture         =   "RolesDetalle.frx":0772
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   225
            Width           =   1005
         End
         Begin VB.CommandButton cmdAddR 
            DisabledPicture =   "RolesDetalle.frx":0B03
            DownPicture     =   "RolesDetalle.frx":0EEC
            Height          =   315
            Left            =   9180
            Picture         =   "RolesDetalle.frx":12F8
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   225
            Width           =   1005
         End
         Begin PVCOMBOLibCtl.PVComboBox cmbListaReportes 
            Height          =   345
            Left            =   600
            TabIndex        =   30
            Top             =   210
            Width           =   5685
            _Version        =   524288
            _cx             =   10028
            _cy             =   609
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
            ColumnHeaderBackColor=   14215660
            SelectedForeColor=   16777215
            SelectedBackColor=   12937777
            AlternateBackColor=   16777215
            ItemLabelStyle  =   1
            ItemLabelType   =   0
            ItemLabelWidth  =   40
            ItemLabelForeColor=   0
            ItemLabelBackColor=   14215660
            ColumnHeaderStyle=   1
            VerticalGridLines=   -1  'True
            HorizontalGridLines=   -1  'True
            ColumnResize    =   -1  'True
            ItemLabelResize =   0   'False
            AllowDBAutoConfig=   -1  'True
            GridLineColor   =   13421772
            List            =   ""
            NullString      =   "[NULL]"
            DropShadow      =   -1  'True
            Text            =   ""
            SortOnColumnHeaderClick=   0   'False
            DropEffect      =   1
            ColumnCount     =   3
            Column0.Heading =   "ID"
            Column0.Width   =   10
            Column0.Alignment=   0
            Column0.Hidden  =   0   'False
            Column0.Name    =   "IdReporte"
            Column0.Format  =   ""
            Column0.Bound   =   -1  'True
            Column0.Locked  =   0   'False
            Column0.HeaderAlignment=   0
            Column1.Heading =   "Reporte"
            Column1.Width   =   400
            Column1.Alignment=   0
            Column1.Hidden  =   0   'False
            Column1.Name    =   "Reporte"
            Column1.Format  =   ""
            Column1.Bound   =   -1  'True
            Column1.Locked  =   0   'False
            Column1.HeaderAlignment=   0
            Column2.Heading =   "Módulo"
            Column2.Width   =   30
            Column2.Alignment=   0
            Column2.Hidden  =   0   'False
            Column2.Name    =   "Modulo"
            Column2.Format  =   ""
            Column2.Bound   =   -1  'True
            Column2.Locked  =   0   'False
            Column2.HeaderAlignment=   0
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
         Begin VB.Label Label2 
            Caption         =   "Item"
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
            TabIndex        =   26
            Top             =   255
            Width           =   510
         End
      End
      Begin VB.Frame Frame4 
         Height          =   675
         Left            =   -74850
         TabIndex        =   17
         Top             =   660
         Width           =   11355
         Begin VB.CommandButton btnQuitarPermiso 
            DisabledPicture =   "RolesDetalle.frx":1704
            DownPicture     =   "RolesDetalle.frx":1A8F
            Height          =   315
            Left            =   10230
            Picture         =   "RolesDetalle.frx":1E22
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   225
            Width           =   1005
         End
         Begin VB.CommandButton btnAgregarPermiso 
            DisabledPicture =   "RolesDetalle.frx":21B3
            DownPicture     =   "RolesDetalle.frx":259C
            Height          =   315
            Left            =   9150
            Picture         =   "RolesDetalle.frx":29A8
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   225
            Width           =   1005
         End
         Begin PVCOMBOLibCtl.PVComboBox cmbIdPermiso 
            Height          =   300
            Left            =   930
            TabIndex        =   20
            Top             =   225
            Width           =   3315
            _Version        =   524288
            _cx             =   5847
            _cy             =   529
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
            ColumnHeaderStyle=   0
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
            Column0.Heading =   "Id"
            Column0.Width   =   20
            Column0.Alignment=   0
            Column0.Hidden  =   0   'False
            Column0.Name    =   "IdPermiso"
            Column0.Format  =   ""
            Column0.Bound   =   -1  'True
            Column0.Locked  =   0   'False
            Column0.HeaderAlignment=   0
            Column1.Heading =   "Descripción"
            Column1.Width   =   200
            Column1.Alignment=   0
            Column1.Hidden  =   0   'False
            Column1.Name    =   "Descripcion"
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
         Begin VB.Label Label1 
            Caption         =   "Permisos"
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
            Left            =   150
            TabIndex        =   21
            Top             =   255
            Width           =   810
         End
      End
      Begin VB.Frame Frame2 
         Height          =   675
         Left            =   120
         TabIndex        =   8
         Top             =   630
         Width           =   11385
         Begin VB.CheckBox chkConsultar 
            Caption         =   "Consultar"
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
            Left            =   6720
            TabIndex        =   15
            Top             =   225
            Width           =   1125
         End
         Begin VB.CheckBox chkEliminar 
            Caption         =   "Eliminar"
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
            Left            =   7905
            TabIndex        =   14
            Top             =   225
            Width           =   1000
         End
         Begin VB.CheckBox chkModificar 
            Caption         =   "Modificar"
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
            Left            =   5565
            TabIndex        =   13
            Top             =   225
            Width           =   1000
         End
         Begin VB.CheckBox chkAgregar 
            Caption         =   "Agregar"
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
            Left            =   4515
            TabIndex        =   12
            Top             =   225
            Width           =   1000
         End
         Begin VB.CommandButton btnAgregarDx 
            DisabledPicture =   "RolesDetalle.frx":2DB4
            DownPicture     =   "RolesDetalle.frx":319D
            Height          =   315
            Left            =   9180
            Picture         =   "RolesDetalle.frx":35A9
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   225
            Width           =   1005
         End
         Begin VB.CommandButton btnQuitarDx 
            DisabledPicture =   "RolesDetalle.frx":39B5
            DownPicture     =   "RolesDetalle.frx":3D40
            Height          =   315
            Left            =   10260
            Picture         =   "RolesDetalle.frx":40D3
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   225
            Width           =   1005
         End
         Begin PVCOMBOLibCtl.PVComboBox cmbIdListItem 
            Height          =   300
            Left            =   930
            TabIndex        =   9
            Top             =   225
            Width           =   3315
            _Version        =   524288
            _cx             =   5847
            _cy             =   529
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
            ColumnCount     =   3
            Column0.Heading =   "Id"
            Column0.Width   =   20
            Column0.Alignment=   0
            Column0.Hidden  =   0   'False
            Column0.Name    =   "IdListItem"
            Column0.Format  =   ""
            Column0.Bound   =   -1  'True
            Column0.Locked  =   0   'False
            Column0.HeaderAlignment=   0
            Column1.Heading =   "SubModulo"
            Column1.Width   =   200
            Column1.Alignment=   0
            Column1.Hidden  =   0   'False
            Column1.Name    =   "SubModulo"
            Column1.Format  =   ""
            Column1.Bound   =   -1  'True
            Column1.Locked  =   0   'False
            Column1.HeaderAlignment=   0
            Column2.Heading =   "Modulo"
            Column2.Width   =   200
            Column2.Alignment=   0
            Column2.Hidden  =   0   'False
            Column2.Name    =   "Modulo"
            Column2.Format  =   ""
            Column2.Bound   =   -1  'True
            Column2.Locked  =   0   'False
            Column2.HeaderAlignment=   0
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
         Begin VB.Label lblIdListItem 
            Caption         =   "Item"
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
            TabIndex        =   16
            Top             =   255
            Width           =   510
         End
      End
      Begin UltraGrid.SSUltraGrid grdPermisos 
         Height          =   4410
         Left            =   120
         TabIndex        =   7
         Top             =   1395
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   7779
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
         Caption         =   "Items asignados"
      End
      Begin UltraGrid.SSUltraGrid grdPermisos2 
         Height          =   4350
         Left            =   -74850
         TabIndex        =   22
         Top             =   1470
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   7673
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
         Caption         =   "Permisos asignados"
      End
      Begin UltraGrid.SSUltraGrid grdReportes 
         Height          =   4380
         Left            =   -74880
         TabIndex        =   27
         Top             =   1425
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   7726
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
         Caption         =   "Reportes asignados"
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   90
      TabIndex        =   5
      Top             =   6810
      Width           =   11670
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "RolesDetalle.frx":4464
         DownPicture     =   "RolesDetalle.frx":48C4
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
         Left            =   4530
         Picture         =   "RolesDetalle.frx":4D39
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "RolesDetalle.frx":51AE
         DownPicture     =   "RolesDetalle.frx":5672
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
         Left            =   6090
         Picture         =   "RolesDetalle.frx":5B5E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   720
      Left            =   105
      TabIndex        =   3
      Top             =   30
      Width           =   11625
      Begin VB.TextBox txtNombre 
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
         MaxLength       =   50
         TabIndex        =   0
         Top             =   270
         Width           =   3360
      End
      Begin VB.Label lblNombre 
         Caption         =   "Nombre"
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
         TabIndex        =   4
         Top             =   270
         Width           =   1005
      End
   End
End
Attribute VB_Name = "RolesDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Registro de Rol
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Teclado As New sighEntidades.Teclado
Dim mo_Formulario As New sighEntidades.Formulario
Dim mo_AdminSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_Roles As New DORol
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdRol As Long

Dim mrs_RolItems As New ADODB.Recordset
Dim mrs_RolPermisos As New ADODB.Recordset
Dim mrs_RolReporte As New ADODB.Recordset
Dim oRsListaReportes As New Recordset

Dim mo_RolItems As New Collection
Dim mo_RolPermisos As New Collection
Dim mo_RolReportes As New Collection

Dim mo_Apariencia As New sighEntidades.GridInfragistic
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Sub CargarComboBoxes()
Dim sSQL As String
Dim sMensaje As String
    Set cmbIdListItem.ListSource = mo_AdminSeguridad.ListItemsSeleccionarTodos()
    Set cmbIdPermiso.ListSource = mo_AdminSeguridad.PermisosSeleccionarTodos()
    '
    Set oRsListaReportes = mo_AdminSeguridad.ListBarReportesSeleccionarTodos
    Set cmbListaReportes.ListSource = oRsListaReportes
End Sub
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
Property Let IdRol(lValue As Long)
   ml_IdRol = lValue
End Property
Property Get IdRol() As Long
   IdRol = ml_IdRol
End Property

Private Sub btnAgregarPermiso_Click()
Dim oCampos() As String

    If Me.cmbIdPermiso.Text = "" Then
        MsgBox "Seleccione un permiso de la lista desplegable", vbInformation, Me.Caption
        Exit Sub
    End If

    oCampos = Split(Me.cmbIdPermiso.List(cmbIdPermiso.ListIndex), "|")
    
    On Error Resume Next
    mrs_RolPermisos.MoveFirst
    Do While Not mrs_RolPermisos.EOF
        If oCampos(0) = mrs_RolPermisos!IdPermiso Then
            MsgBox "El permiso ya fue asignado", vbExclamation, Me.Caption
            Exit Sub
        End If
        mrs_RolPermisos.MoveNext
    Loop
    
    With mrs_RolPermisos
        .AddNew
        .Fields!IdPermiso = oCampos(0)
        .Fields!descripcion = oCampos(1)
    End With


End Sub

Private Sub btnQuitarPermiso_Click()
    On Error Resume Next
    With mrs_RolPermisos
        If Not .EOF And Not .BOF Then
           .Delete
           .Update
        End If
    End With
End Sub



Private Sub chkAgregar_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkAgregar
    AdministrarKeyPreview KeyCode
End Sub
Private Sub chkModificar_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkModificar
    AdministrarKeyPreview KeyCode
End Sub
Private Sub chkConsultar_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkConsultar
    AdministrarKeyPreview KeyCode
End Sub
Private Sub chkEliminar_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkEliminar
    AdministrarKeyPreview KeyCode
End Sub

Private Sub chkNingunReporte_Click()
    If chkNingunReporte.Value = 1 Then
       If mrs_RolReporte.RecordCount > 0 Then
          mrs_RolReporte.MoveFirst
          Do While Not mrs_RolReporte.EOF
             mrs_RolReporte.Delete
             mrs_RolReporte.Update
             mrs_RolReporte.MoveNext
          Loop
       End If
    End If
End Sub

Private Sub chkTodosReportes_Click()
   If chkTodosReportes.Value = 1 Then
        Dim lnCorr As Integer
        Dim oCampos() As String
        Dim lbNuevo As Boolean
        For lnCorr = 1 To cmbListaReportes.ListCount
             oCampos = Split(Me.cmbListaReportes.List(lnCorr - 1), "|")
             lbNuevo = True
             If mrs_RolReporte.RecordCount > 0 Then
                mrs_RolReporte.MoveFirst
                mrs_RolReporte.Find "idReporte=" & oCampos(0)
                If Not mrs_RolReporte.EOF Then
                   lbNuevo = False
                End If
             End If
             If lbNuevo = True Then
                mrs_RolReporte.AddNew
                mrs_RolReporte.Fields!idReporte = oCampos(0)
                mrs_RolReporte.Fields!reporte = oCampos(1)
                mrs_RolReporte.Fields!Modulo = oCampos(2)
                mrs_RolReporte.Fields!tieneAcceso = 1
                mrs_RolReporte.Update
             End If
        Next
   End If
End Sub

Private Sub cmbDelR_Click()
    On Error Resume Next
    With mrs_RolReporte
        If Not .EOF And Not .BOF Then
           .Delete
           .Update
        End If
    End With
End Sub

Private Sub cmbIdListItem_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdListItem
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdListItem_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub cmbIdPermiso_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdListItem
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdPermiso_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub



Private Sub cmdAddR_Click()
    Dim lbNuevo As Boolean
    Dim oCampos() As String
    If Me.cmbListaReportes.Text = "" Then
        MsgBox "Seleccione un item de la lista desplegable", vbInformation, Me.Caption
        Exit Sub
    End If
    oCampos = Split(Me.cmbListaReportes.List(cmbListaReportes.ListIndex), "|")
    lbNuevo = True
    If mrs_RolReporte.RecordCount > 0 Then
       mrs_RolReporte.MoveFirst
       mrs_RolReporte.Find "idReporte=" & oCampos(0)
       If Not mrs_RolReporte.EOF Then
          lbNuevo = False
       End If
    End If
    If lbNuevo = True Then
       mrs_RolReporte.AddNew
       mrs_RolReporte.Fields!idReporte = oCampos(0)
       mrs_RolReporte.Fields!reporte = oCampos(1)
       mrs_RolReporte.Fields!Modulo = oCampos(2)
       mrs_RolReporte.Fields!tieneAcceso = 1
       mrs_RolReporte.Update
    Else
       MsgBox "El item ya fue asignado", vbExclamation, Me.Caption
    End If
End Sub



Private Sub grdPermisos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdPermisos.Bands(0).Columns("IdListItem").Hidden = True
    
    grdPermisos.Bands(0).Columns("SubModulo").Header.Caption = "SubModulo"
    grdPermisos.Bands(0).Columns("SubModulo").Width = 2500
    
    grdPermisos.Bands(0).Columns("Modulo").Header.Caption = "Modulo"
    grdPermisos.Bands(0).Columns("Modulo").Width = 2500
    
    grdPermisos.Bands(0).Columns("Agregar").Header.Caption = "Agregar"
    grdPermisos.Bands(0).Columns("Agregar").Width = 1000

    grdPermisos.Bands(0).Columns("Modificar").Header.Caption = "Modificar"
    grdPermisos.Bands(0).Columns("Modificar").Width = 1000

    grdPermisos.Bands(0).Columns("Consultar").Header.Caption = "Consultar"
    grdPermisos.Bands(0).Columns("Consultar").Width = 1000

    grdPermisos.Bands(0).Columns("Eliminar").Header.Caption = "Eliminar"
    grdPermisos.Bands(0).Columns("Eliminar").Width = 1000


End Sub

Private Sub grdPermisos2_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdPermisos2.Bands(0).Columns("IdPermiso").Hidden = True
    
    grdPermisos2.Bands(0).Columns("Descripcion").Header.Caption = "Descripcion"
    grdPermisos2.Bands(0).Columns("Descripcion").Width = 7500
    
End Sub

Private Sub grdReportes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdReportes.Bands(0).Columns("Reporte").Header.Caption = "Reporte"
    grdReportes.Bands(0).Columns("Reporte").Width = 7500
    
    grdReportes.Bands(0).Columns("Modulo").Header.Caption = "Módulo"
    grdReportes.Bands(0).Columns("Modulo").Width = 2500
    grdReportes.Bands(0).Columns("idReporte").Hidden = True
End Sub



Private Sub PVListBox1_Click()

End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNombre
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtNombre_LostFocus()
   mo_Formulario.MarcarComoVacio txtNombre
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Roles
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()

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
        Me.Frame1.Enabled = False
        Me.Frame2.Enabled = False
        Me.btnAceptar.Enabled = False
     Case sghEliminar
        Me.Frame1.Enabled = False
        Me.Frame2.Enabled = False
 End Select

End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Roles
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       
    GenerarRecordsetTemporal
        
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Roles"
       Case sghModificar
           Me.Caption = "Modificar Roles"
       Case sghConsultar
           Me.Caption = "Consultar Roles"
       Case sghEliminar
           Me.Caption = "Eliminar Roles"
       End Select

       CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Roles
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
       End Select
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   AdministrarKeyPreview KeyCode
End Sub

Private Sub btnAceptar_Click()
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If AgregarDatos() Then
                   MsgBox " Los datos se agregaron exitosamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminSeguridad.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If ModificarDatos() Then
                   MsgBox " Los datos se modificaron exitosamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminSeguridad.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                   MsgBox " Los datos se eliminaron exitosamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_AdminSeguridad.MensajeError, vbExclamation, Me.Caption
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
   If Me.txtNombre.Text = "" Then
       sMensaje = sMensaje + "Ingrese el nombre del rol" + Chr(13)
   End If
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   Dim lnCorr As Integer
   Dim oCampos() As String
   Dim lbNuevo As Boolean
   Set Me.grdReportes.DataSource = Nothing
   For lnCorr = 1 To cmbListaReportes.ListCount
        oCampos = Split(Me.cmbListaReportes.List(lnCorr - 1), "|")
        lbNuevo = True
        If mrs_RolReporte.RecordCount > 0 Then
           mrs_RolReporte.MoveFirst
           mrs_RolReporte.Find "idReporte=" & oCampos(0)
           If Not mrs_RolReporte.EOF Then
              lbNuevo = False
           End If
        End If
        If lbNuevo = True Then
           mrs_RolReporte.AddNew
           mrs_RolReporte.Fields!idReporte = oCampos(0)
           mrs_RolReporte.Fields!reporte = oCampos(1)
           mrs_RolReporte.Fields!Modulo = oCampos(2)
           mrs_RolReporte.Fields!tieneAcceso = 0
           mrs_RolReporte.Update
        End If
   Next
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
   ValidarReglas = False
   ValidarReglas = True
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla Roles
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()

   With mo_Roles
           .IdRol = Me.IdRol
           .nombre = Me.txtNombre.Text
           .IdUsuarioAuditoria = Me.idUsuario
   End With
    
    CargarRolItemsAlObjetoDatos mo_RolItems
    CargarPermisosAlObjetoDatos mo_RolPermisos
   
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   AgregarDatos = mo_AdminSeguridad.RolesAgregar(mo_Roles, mo_RolItems, mo_RolPermisos, mrs_RolReporte, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtNombre.Text)
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   ModificarDatos = mo_AdminSeguridad.RolesModificar(mo_Roles, mo_RolItems, mo_RolPermisos, mrs_RolReporte, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtNombre.Text)

End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   EliminarDatos = mo_AdminSeguridad.RolesEliminar(mo_Roles, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtNombre.Text)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Roles
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlosControles()

       Set mo_Roles = mo_AdminSeguridad.RolesSeleccionarPorId(Me.IdRol)
        If mo_AdminSeguridad.MensajeError <> "" Then
            MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminSeguridad.MensajeError, vbInformation, Me.Caption"
            mb_ExistenDatos = False
            Exit Sub
        End If
        
        If Not mo_Roles Is Nothing Then
           With mo_Roles
                Me.IdRol = .IdRol
                Me.txtNombre.Text = .nombre
                mb_ExistenDatos = True
           End With
           CargarDatosDeRolItems
           
        Else
           mb_ExistenDatos = False
           Exit Sub
        End If
        
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Roles
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()
           Me.IdRol = 0
           Me.txtNombre.Text = ""
End Sub
Sub CargarDatosDeRolItems()
Dim rsRolItems As New Recordset
Dim rsPermisos As New Recordset
Dim rsRolReportes As New Recordset
Dim lcSql As String

    Set rsRolItems = mo_AdminSeguridad.RolesItemsSeleccionarPorRol(ml_IdRol)
    Do While Not rsRolItems.EOF
        With mrs_RolItems
            .AddNew
            .Fields!IdListItem = rsRolItems!IdListItem
            .Fields!SubModulo = rsRolItems!SubModulo
            .Fields!Modulo = rsRolItems!Modulo
            .Fields!Agregar = rsRolItems!Agregar
            .Fields!Modificar = rsRolItems!Modificar
            .Fields!Consultar = rsRolItems!Consultar
            .Fields!Eliminar = rsRolItems!Eliminar
        End With
        rsRolItems.MoveNext
    Loop
    mo_Apariencia.ConfigurarFilasBiColores Me.grdPermisos, sighEntidades.GrillaConFilasBicolor
    
    Set rsPermisos = mo_AdminSeguridad.RolesPermisosSeleccionarPorRol(ml_IdRol)
    Do While Not rsPermisos.EOF
        With mrs_RolPermisos
            .AddNew
            .Fields!IdPermiso = rsPermisos!IdPermiso
            .Fields!descripcion = rsPermisos!descripcion
        End With
        rsPermisos.MoveNext
    Loop
    mo_Apariencia.ConfigurarFilasBiColores Me.grdPermisos2, sighEntidades.GrillaConFilasBicolor
    '
    Set rsRolReportes = mo_AdminSeguridad.RolesReportesSeleccionarXrol(ml_IdRol)
    If rsRolReportes.RecordCount > 0 Then
       rsRolReportes.MoveFirst
       Do While Not rsRolReportes.EOF
          With mrs_RolReporte
              .AddNew
              .Fields!idReporte = rsRolReportes.Fields!idReporte
              .Fields!reporte = rsRolReportes.Fields!reporte
              .Fields!Modulo = rsRolReportes.Fields!Modulo
              .Fields!tieneAcceso = rsRolReportes.Fields!tieneAcceso
              .Update
          End With
          rsRolReportes.MoveNext
       Loop
    End If
    mo_Apariencia.ConfigurarFilasBiColores Me.grdReportes, sighEntidades.GrillaConFilasBicolor
    
End Sub

Sub CargarRolItemsAlObjetoDatos(oRolItems As Collection)
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LOS ExamenS
    '---------------------------------------------------------------------------------
    Dim oRolItem As DORolItem
    
    If Not (mrs_RolItems.BOF And mrs_RolItems.EOF) Then
        mrs_RolItems.MoveFirst
        Do While Not mrs_RolItems.EOF
            Set oRolItem = New DORolItem
            oRolItem.IdRolItem = 0
            oRolItem.IdRol = ml_IdRol
            oRolItem.IdListItem = mrs_RolItems!IdListItem
            oRolItem.Agregar = mrs_RolItems!Agregar
            oRolItem.Modificar = mrs_RolItems!Modificar
            oRolItem.Consultar = "" & mrs_RolItems!Consultar
            oRolItem.Eliminar = mrs_RolItems!Eliminar
            oRolItem.IdUsuarioAuditoria = ml_idUsuario
            oRolItems.Add oRolItem
            
            mrs_RolItems.MoveNext
        Loop
    End If
    
End Sub

Sub CargarPermisosAlObjetoDatos(oRolPermisos As Collection)
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LOS ExamenS
    '---------------------------------------------------------------------------------
    Dim oRolPermiso As DORolPermiso
    
    If Not (mrs_RolPermisos.BOF And mrs_RolPermisos.EOF) Then
        mrs_RolPermisos.MoveFirst
        Do While Not mrs_RolPermisos.EOF
            Set oRolPermiso = New DORolPermiso
            oRolPermiso.IdRolPermiso = 0
            oRolPermiso.IdRol = ml_IdRol
            oRolPermiso.IdPermiso = mrs_RolPermisos!IdPermiso
            oRolPermiso.IdUsuarioAuditoria = ml_idUsuario
            oRolPermisos.Add oRolPermiso
            mrs_RolPermisos.MoveNext
        Loop
    End If
    
End Sub


Sub GenerarRecordsetTemporal()
    
    With mrs_RolItems
          .Fields.Append "IdListItem", adInteger, 4, adFldIsNullable
          .Fields.Append "SubModulo", adVarChar, 100, adFldIsNullable
          .Fields.Append "Modulo", adVarChar, 100, adFldIsNullable
          .Fields.Append "Agregar", adBoolean
          .Fields.Append "Modificar", adBoolean
          .Fields.Append "Consultar", adBoolean
          .Fields.Append "Eliminar", adBoolean
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
    Set Me.grdPermisos.DataSource = mrs_RolItems
    mo_Apariencia.ConfigurarFilasBiColores Me.grdPermisos, sighEntidades.GrillaConFilasBicolor
    
    With mrs_RolPermisos
          .Fields.Append "IdPermiso", adInteger, 4, adFldIsNullable
          .Fields.Append "Descripcion", adVarChar, 200, adFldIsNullable
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
    Set Me.grdPermisos2.DataSource = mrs_RolPermisos
    mo_Apariencia.ConfigurarFilasBiColores Me.grdPermisos2, sighEntidades.GrillaConFilasBicolor
    
    With mrs_RolReporte
          .Fields.Append "IdReporte", adInteger, 4, adFldIsNullable
          .Fields.Append "Reporte", adVarChar, 255, adFldIsNullable
          .Fields.Append "Modulo", adVarChar, 100, adFldIsNullable
          .Fields.Append "TieneAcceso", adInteger, 4, adFldIsNullable
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
    Set Me.grdReportes.DataSource = mrs_RolReporte
    mo_Apariencia.ConfigurarFilasBiColores Me.grdReportes, sighEntidades.GrillaConFilasBicolor
End Sub

Private Sub btnAgregarDx_Click()
Dim oCampos() As String

    If Me.cmbIdListItem.Text = "" Then
        MsgBox "Seleccione un item de la lista desplegable", vbInformation, Me.Caption
        Exit Sub
    End If

    oCampos = Split(Me.cmbIdListItem.List(cmbIdListItem.ListIndex), "|")
    
    On Error Resume Next
    mrs_RolItems.MoveFirst
    Do While Not mrs_RolItems.EOF
        If oCampos(0) = mrs_RolItems!IdListItem Then
            MsgBox "El item ya fue asignado", vbExclamation, Me.Caption
            Exit Sub
        End If
        mrs_RolItems.MoveNext
    Loop
    
    With mrs_RolItems
        .AddNew
        .Fields!IdListItem = oCampos(0)
        .Fields!SubModulo = oCampos(1)
        .Fields!Modulo = oCampos(2)
        .Fields!Agregar = Me.chkAgregar.Value
        .Fields!Modificar = Me.chkModificar.Value
        .Fields!Consultar = Me.chkConsultar.Value
        .Fields!Eliminar = Me.chkEliminar.Value
    End With

End Sub

Private Sub btnQuitarDx_Click()
    On Error Resume Next
    With mrs_RolItems
        If Not .EOF And Not .BOF Then
           .Delete
           .Update
        End If
    End With
End Sub


