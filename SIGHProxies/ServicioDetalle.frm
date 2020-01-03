VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form ServicioDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13410
   Icon            =   "ServicioDetalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   13410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraServicios 
      Caption         =   "Datos para Hospitalización/Emergencia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2970
      Left            =   165
      TabIndex        =   34
      Top             =   4500
      Width           =   6075
      Begin VB.TextBox txtEmergCorrelativo 
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
         Left            =   1575
         MaxLength       =   10
         TabIndex        =   76
         Top             =   1455
         Width           =   1050
      End
      Begin VB.CheckBox chkFichaNacimiento 
         Caption         =   "Usa ficha NACIMIENTO"
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
         Left            =   180
         TabIndex        =   75
         Top             =   2505
         Width           =   5385
      End
      Begin VB.CheckBox chkFuaAdmisionEmerg 
         Caption         =   "Se emite Formato FUA desde Admisión Emergencia"
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
         Left            =   195
         TabIndex        =   15
         Top             =   2160
         Width           =   5385
      End
      Begin VB.CheckBox chkEsObsEmerg 
         Caption         =   "El Servicio es 'Observación de Emergencia'"
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
         Left            =   195
         TabIndex        =   14
         Top             =   1830
         Width           =   4005
      End
      Begin VB.TextBox txtUbicacionSEM 
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
         Left            =   1575
         MaxLength       =   5
         TabIndex        =   13
         Top             =   1080
         Width           =   1050
      End
      Begin VB.TextBox txtCodigoSEM 
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
         Left            =   1575
         MaxLength       =   6
         TabIndex        =   12
         Top             =   690
         Width           =   1050
      End
      Begin VB.CommandButton cmbBuscaEstancia 
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
         Left            =   2700
         TabIndex        =   10
         Top             =   300
         Width           =   435
      End
      Begin VB.TextBox txtProductoEstancia 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3180
         MaxLength       =   50
         TabIndex        =   11
         Top             =   300
         Width           =   2850
      End
      Begin VB.TextBox txtCodProductoEstancia 
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
         Left            =   1575
         TabIndex        =   9
         Top             =   300
         Width           =   1050
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "(usado en LIBRO DE EMERGENCIA)"
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
         Left            =   2730
         TabIndex        =   78
         Top             =   1530
         Width           =   2880
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Correlativo N°"
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
         Left            =   165
         TabIndex        =   77
         Top             =   1470
         Width           =   1125
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "(Para  Exportar datos al Sistema SEM)"
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
         Left            =   2730
         TabIndex        =   45
         Top             =   1140
         Width           =   3090
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "(Para  Exportar datos al Sistema SEM)"
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
         Left            =   2730
         TabIndex        =   44
         Top             =   750
         Width           =   3090
      End
      Begin VB.Label Label7 
         Caption         =   "Ubicación"
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
         Left            =   180
         TabIndex        =   43
         Top             =   1080
         Width           =   1395
      End
      Begin VB.Label Label9 
         Caption         =   "Código Servicio"
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
         Left            =   180
         TabIndex        =   42
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label Label14 
         Caption         =   "Estancia por día"
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
         Left            =   180
         TabIndex        =   35
         Top             =   330
         Width           =   1305
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   30
      TabIndex        =   28
      Top             =   7545
      Width           =   13335
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "ServicioDetalle.frx":08CA
         DownPicture     =   "ServicioDetalle.frx":0D8E
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
         Left            =   6795
         Picture         =   "ServicioDetalle.frx":127A
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   255
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "ServicioDetalle.frx":1766
         DownPicture     =   "ServicioDetalle.frx":1BC6
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
         Left            =   5250
         Picture         =   "ServicioDetalle.frx":203B
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   255
         Width           =   1365
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
      Height          =   7485
      Left            =   45
      TabIndex        =   27
      Top             =   30
      Width           =   13335
      Begin VB.Frame Frame 
         Caption         =   "SIS"
         Height          =   1050
         Left            =   105
         TabIndex        =   63
         Top             =   3345
         Width           =   6090
         Begin VB.ComboBox cmdFUAanexo 
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
            ItemData        =   "ServicioDetalle.frx":24B0
            Left            =   1575
            List            =   "ServicioDetalle.frx":24BA
            Style           =   2  'Dropdown List
            TabIndex        =   65
            Top             =   555
            Width           =   4395
         End
         Begin VB.CheckBox chkUsaFormatoFUA 
            Caption         =   "Se usa Formato FUA"
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
            Left            =   165
            TabIndex        =   64
            Top             =   225
            Width           =   2130
         End
         Begin VB.Label lblFuaAnexo 
            AutoSize        =   -1  'True
            Caption         =   "Anexo 2015"
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
            Left            =   495
            TabIndex        =   66
            Top             =   600
            Width           =   1005
         End
      End
      Begin VB.Frame Frame8 
         Height          =   3135
         Left            =   105
         TabIndex        =   36
         Top             =   150
         Width           =   6075
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
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   4
            Top             =   1590
            Width           =   4275
         End
         Begin VB.TextBox txtCodigo 
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
            MaxLength       =   6
            TabIndex        =   3
            Top             =   1230
            Width           =   1050
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
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   150
            Width           =   4305
         End
         Begin VB.ComboBox cmbIdDepartamento 
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
            TabIndex        =   1
            Top             =   510
            Width           =   4305
         End
         Begin VB.ComboBox cmbIdEspecialidad 
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
            TabIndex        =   2
            Top             =   870
            Width           =   4305
         End
         Begin VB.CheckBox chkPuntoCarga 
            Caption         =   "El Servicio es un PUNTO DE CARGA"
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
            TabIndex        =   7
            Top             =   2670
            Width           =   3285
         End
         Begin VB.CheckBox chkEstado 
            Alignment       =   1  'Right Justify
            Caption         =   "Habilitado"
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
            Left            =   4800
            TabIndex        =   8
            Top             =   2655
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin PVCOMBOLibCtl.PVComboBox cmbUPSsusalud 
            Height          =   330
            Left            =   1650
            TabIndex        =   6
            Top             =   2295
            Width           =   1545
            _Version        =   524288
            _cx             =   2725
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
            Column0.Heading =   "Descripción"
            Column0.Width   =   200
            Column0.Alignment=   0
            Column0.Hidden  =   0   'False
            Column0.Name    =   "descripcion"
            Column0.Format  =   ""
            Column0.Bound   =   -1  'True
            Column0.Locked  =   0   'False
            Column0.HeaderAlignment=   0
            Column1.Heading =   "Código"
            Column1.Width   =   60
            Column1.Alignment=   0
            Column1.Hidden  =   0   'False
            Column1.Name    =   "codigo"
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
         Begin PVCOMBOLibCtl.PVComboBox cmbUPSfua 
            Height          =   330
            Left            =   1650
            TabIndex        =   5
            Top             =   1920
            Width           =   1545
            _Version        =   524288
            _cx             =   2725
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
            Column0.Heading =   "Descripción"
            Column0.Width   =   200
            Column0.Alignment=   0
            Column0.Hidden  =   0   'False
            Column0.Name    =   "descripcion"
            Column0.Format  =   ""
            Column0.Bound   =   -1  'True
            Column0.Locked  =   0   'False
            Column0.HeaderAlignment=   0
            Column1.Heading =   "Código"
            Column1.Width   =   60
            Column1.Alignment=   0
            Column1.Hidden  =   0   'False
            Column1.Name    =   "ups"
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
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Código Servicio"
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
            TabIndex        =   62
            Top             =   2340
            Width           =   1230
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "(UPS para SuSalud)"
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
            Left            =   3180
            TabIndex        =   61
            Top             =   2340
            Width           =   1590
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Código Servicio"
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
            TabIndex        =   60
            Top             =   1995
            Width           =   1230
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "(UPS para FUA)"
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
            Left            =   3195
            TabIndex        =   59
            Top             =   1995
            Width           =   1275
         End
         Begin VB.Label Label8 
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
            Left            =   120
            TabIndex        =   41
            Top             =   900
            Width           =   1395
         End
         Begin VB.Label Label6 
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
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   210
            Width           =   1395
         End
         Begin VB.Label Label1 
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
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   1590
            Width           =   1395
         End
         Begin VB.Label Label2 
            Caption         =   "Departamento"
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
            TabIndex        =   38
            Top             =   540
            Width           =   1395
         End
         Begin VB.Label Label3 
            Caption         =   "Código"
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
            TabIndex        =   37
            Top             =   1230
            Width           =   1395
         End
      End
      Begin VB.Frame FraConsultorio 
         Caption         =   "Datos para Consultorio Externo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6120
         Left            =   6270
         TabIndex        =   32
         Top             =   1320
         Width           =   7005
         Begin VB.CheckBox chkNoUsaMTcelular 
            Alignment       =   1  'Right Justify
            Caption         =   "No se usa MENSAJE TEXTO a Celular"
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
            Left            =   3495
            TabIndex        =   79
            Top             =   930
            Width           =   3300
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
            Height          =   360
            Left            =   4920
            Picture         =   "ServicioDetalle.frx":24D2
            Style           =   1  'Graphical
            TabIndex        =   74
            ToolTipText     =   "Lista de ACTIVIDADES HIS de acuerdo a la EDAD, PESO  y UPS"
            Top             =   3555
            Width           =   480
         End
         Begin VB.TextBox txtMaxCitasSISadelantadas 
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
            Left            =   4050
            MaxLength       =   3
            TabIndex        =   53
            Text            =   "90"
            Top             =   4575
            Width           =   735
         End
         Begin VB.TextBox txtMaxCitasSisHoy 
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
            Left            =   4050
            MaxLength       =   3
            TabIndex        =   54
            Text            =   "90"
            Top             =   4920
            Width           =   735
         End
         Begin VB.TextBox txtMaxCitasAdicionales 
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
            Left            =   4050
            MaxLength       =   3
            TabIndex        =   52
            Text            =   "90"
            Top             =   4230
            Width           =   735
         End
         Begin VB.TextBox txtMaxCitasAdelantadas 
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
            Left            =   4050
            MaxLength       =   3
            TabIndex        =   51
            Text            =   "90"
            Top             =   3885
            Width           =   735
         End
         Begin VB.Frame fraConsultorios 
            Caption         =   "UPS adicionales en el Consultorio"
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
            Height          =   1425
            Left            =   150
            TabIndex        =   48
            Top             =   2115
            Width           =   6675
            Begin VB.ComboBox cmbConsultorios 
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
               ItemData        =   "ServicioDetalle.frx":2A5C
               Left            =   120
               List            =   "ServicioDetalle.frx":2A69
               TabIndex        =   58
               Top             =   210
               Width           =   5610
            End
            Begin VB.CommandButton btnAgregar 
               DisabledPicture =   "ServicioDetalle.frx":2A99
               DownPicture     =   "ServicioDetalle.frx":2E82
               Height          =   315
               Left            =   5760
               Picture         =   "ServicioDetalle.frx":328E
               Style           =   1  'Graphical
               TabIndex        =   57
               Top             =   210
               Width           =   825
            End
            Begin UltraGrid.SSUltraGrid grdConsultorios 
               Height          =   810
               Left            =   120
               TabIndex        =   49
               Top             =   570
               Width           =   6480
               _ExtentX        =   11430
               _ExtentY        =   1429
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
               Caption         =   "UPS"
            End
         End
         Begin VB.Frame Frame4 
            Height          =   870
            Left            =   150
            TabIndex        =   47
            Top             =   1215
            Width           =   6675
            Begin VB.CheckBox chkUsaModuloNinoSano 
               Caption         =   "El CONSULTORIO usa el MODULO DE NIÑO SANO"
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
               Left            =   150
               TabIndex        =   23
               Top             =   165
               Width           =   5835
            End
            Begin VB.CheckBox chkUsaModuloMaterno 
               Caption         =   "El CONSULTORIO usa el MODULO MATERNO"
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
               Left            =   150
               TabIndex        =   24
               Top             =   495
               Width           =   5835
            End
         End
         Begin VB.CheckBox chkEnGalenHos 
            Caption         =   "Emite Formato FUA desde CITAS"
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
            Left            =   150
            TabIndex        =   22
            Top             =   930
            Width           =   3090
         End
         Begin VB.CheckBox chkTriaje 
            Caption         =   "El CONSULTORIO necesita TRIAJE antes de registrar ATENCION"
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
            Left            =   150
            TabIndex        =   21
            Top             =   606
            Width           =   5835
         End
         Begin VB.CheckBox chkCostoCero 
            Caption         =   "El CONSULTORIO no cobra por 'atención Consultorio ' a  Particulares"
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
            Left            =   150
            TabIndex        =   20
            Top             =   270
            Width           =   5955
         End
         Begin PVCOMBOLibCtl.PVComboBox cmdCodigoHIS 
            Height          =   330
            Left            =   1455
            TabIndex        =   50
            Top             =   3570
            Width           =   1545
            _Version        =   524288
            _cx             =   2725
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
            ColumnCount     =   3
            Column0.Heading =   "Descripción"
            Column0.Width   =   200
            Column0.Alignment=   0
            Column0.Hidden  =   0   'False
            Column0.Name    =   "descripcion"
            Column0.Format  =   ""
            Column0.Bound   =   -1  'True
            Column0.Locked  =   0   'False
            Column0.HeaderAlignment=   0
            Column1.Heading =   "UPShis"
            Column1.Width   =   60
            Column1.Alignment=   0
            Column1.Hidden  =   0   'False
            Column1.Name    =   "UPShis"
            Column1.Format  =   ""
            Column1.Bound   =   -1  'True
            Column1.Locked  =   0   'False
            Column1.HeaderAlignment=   0
            Column2.Heading =   "IdUps"
            Column2.Width   =   200
            Column2.Alignment=   0
            Column2.Hidden  =   0   'False
            Column2.Name    =   "IdUps"
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
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "(hoy)"
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
            Left            =   4875
            TabIndex        =   73
            Top             =   4950
            Width           =   450
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Máximo CUPOS para CITAS SIS ADELANTADAS"
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
            TabIndex        =   72
            Top             =   4605
            Width           =   3855
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Máximo CUPOS para CITAS SIS"
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
            Top             =   4935
            Width           =   2535
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "(Citas mayores a hoy)"
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
            Left            =   4875
            TabIndex        =   70
            Top             =   4605
            Width           =   1770
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "(Citas mayores a hoy)"
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
            Left            =   4875
            TabIndex        =   69
            Top             =   3915
            Width           =   1770
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Máximo CUPOS para CITAS ADICIONALES"
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
            TabIndex        =   68
            Top             =   4245
            Width           =   3405
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Máximo CUPOS para CITAS ADELANTADAS"
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
            TabIndex        =   67
            Top             =   3915
            Width           =   3525
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "(UPS del Sistema HIS)"
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
            Left            =   3045
            TabIndex        =   56
            Top             =   3600
            Width           =   1815
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Código Servicio"
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
            TabIndex        =   55
            Top             =   3600
            Width           =   1230
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Sólo atiende a  personas"
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
         Left            =   6270
         TabIndex        =   29
         Top             =   150
         Width           =   6975
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
            ItemData        =   "ServicioDetalle.frx":369A
            Left            =   1290
            List            =   "ServicioDetalle.frx":369C
            TabIndex        =   17
            Top             =   660
            Width           =   2100
         End
         Begin VB.TextBox txtEdadMin 
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
            Left            =   4350
            MaxLength       =   3
            TabIndex        =   18
            Top             =   690
            Width           =   735
         End
         Begin VB.TextBox txtEdadMax 
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
            Left            =   5880
            TabIndex        =   19
            Top             =   660
            Width           =   885
         End
         Begin VB.ComboBox cmbTipoSexo 
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
            ItemData        =   "ServicioDetalle.frx":369E
            Left            =   1290
            List            =   "ServicioDetalle.frx":36AB
            TabIndex        =   16
            Top             =   270
            Width           =   5475
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "desde"
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
            Left            =   3810
            TabIndex        =   46
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "hasta "
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
            Left            =   5340
            TabIndex        =   33
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Rango Edad"
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
            TabIndex        =   31
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Sexo"
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
            TabIndex        =   30
            Top             =   300
            Width           =   825
         End
      End
   End
End
Attribute VB_Name = "ServicioDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Consultorios y Servicios Hospitalarios
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_Facturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim ml_idServicio As Long
Dim mo_Servicios As New doServicio
Dim ml_idTipoServicio As Long
Dim mo_cmbIdDepartamento As New sighentidades.ListaDespleglable
Dim mo_cmbIdTipoServicio As New sighentidades.ListaDespleglable
Dim mo_cmbIdEspecialidad As New sighentidades.ListaDespleglable
Dim mo_cmbIdTipoEdad As New sighentidades.ListaDespleglable
Dim mo_cmbConsultorios As New sighentidades.ListaDespleglable
Dim oRsConsultoriosMenosElactual As New Recordset
Dim mRs_ConsultoriosAtencSimultanea As New Recordset
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim lcMensajeLicencia As String

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
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
Property Let idServicio(lValue As Long)
   ml_idServicio = lValue
End Property
Property Get idServicio() As Long
   idServicio = ml_idServicio
End Property
Property Let idTipoServicio(lValue As Long)
   ml_idTipoServicio = lValue
End Property
Property Get idTipoServicio() As Long
   idTipoServicio = ml_idTipoServicio
End Property





Private Sub chkCostoCero_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkCostoCero
    AdministrarKeyPreview KeyCode
End Sub

Private Sub chkEnGalenHos_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkEnGalenHos
    AdministrarKeyPreview KeyCode
End Sub

Private Sub chkEsObsEmerg_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkEsObsEmerg
    AdministrarKeyPreview KeyCode
End Sub

Private Sub chkEstado_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkEstado
    AdministrarKeyPreview KeyCode
End Sub
'mgaray20140926
Private Sub chkFuaAdmisionEmerg_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkFuaAdmisionEmerg
    AdministrarKeyPreview KeyCode
End Sub

Private Sub chkPuntoCarga_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkPuntoCarga
    AdministrarKeyPreview KeyCode
End Sub

Private Sub chkTriaje_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkTriaje
    AdministrarKeyPreview KeyCode
End Sub

Private Sub chkUsaFormatoFUA_Click()
    If chkUsaFormatoFUA.Value = 1 And lcBuscaParametro.SeleccionaFilaParametro(358) = "B" Then
       lblFuaAnexo.Visible = True
       cmdFUAanexo.Visible = True
    Else
       lblFuaAnexo.Visible = False
       cmdFUAanexo.Visible = False
    End If
    If chkUsaFormatoFUA.Value = 1 Then
       chkFuaAdmisionEmerg.Visible = True
       chkEnGalenHos.Visible = True
    Else
       chkFuaAdmisionEmerg.Visible = False
       chkEnGalenHos.Visible = False
       chkFuaAdmisionEmerg.Value = 0
       chkEnGalenHos.Value = 0
    End If
End Sub

Private Sub chkUsaFormatoFUA_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkUsaFormatoFUA
    AdministrarKeyPreview KeyCode
End Sub

Private Sub chkUsaModuloMaterno_Click()
    If chkUsaModuloMaterno.Value = 1 Then
       chkUsaModuloNinoSano.Value = 0
    End If
End Sub

Private Sub chkUsaModuloMaterno_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkUsaModuloMaterno
    AdministrarKeyPreview KeyCode
End Sub

Private Sub chkUsaModuloNinoSano_Click()
    If chkUsaModuloNinoSano.Value = 1 Then
       chkUsaModuloMaterno.Value = 0
    End If
End Sub

Private Sub chkUsaModuloNinoSano_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkUsaModuloNinoSano
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbBuscaEstancia_Click()
        Dim oFrm As New SIGHNegocios.BuscaServicio
        oFrm.MostrarFormulario
        If oFrm.IdRegistroSeleccionado <> 0 Then
            Me.txtCodProductoEstancia.Tag = CStr(oFrm.IdRegistroSeleccionado)
            Call ObtenerNombreServicio(oFrm.IdRegistroSeleccionado, Me.txtCodProductoEstancia, Me.txtProductoEstancia)
        End If
        Set oFrm = Nothing
End Sub

Sub ObtenerNombreServicio(idServicio As Long, txtCode As TextBox, txtName As TextBox)
    Dim dOServ As New DOCatalogoServicio
    Set dOServ = mo_Facturacion.CatalogoServiciosSeleccionarPorId(idServicio)
    If Not dOServ Is Nothing Then
        txtCode.Text = dOServ.codigo
        txtName.Text = dOServ.nombre
    End If
    Set dOServ = Nothing
End Sub



Private Sub cmbIdDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdDepartamento
AdministrarKeyPreview KeyCode
End Sub

Sub CargarComboBoxes()
Dim sSQL As String
Dim sMensaje As String
       
       mo_cmbIdDepartamento.BoundColumn = "IdDepartamento"
       mo_cmbIdDepartamento.ListField = "DescripcionLarga"
       Set mo_cmbIdDepartamento.RowSource = mo_AdminServiciosHosp.DepartamentosSeleccionarTodos()
       sMensaje = sMensaje + mo_AdminServiciosHosp.MensajeError
       
       mo_cmbIdTipoServicio.BoundColumn = "IdTipoServicio"
       mo_cmbIdTipoServicio.ListField = "DescripcionLarga"
       Set mo_cmbIdTipoServicio.RowSource = mo_AdminServiciosHosp.TiposServicioSeleccionarTodos()
       sMensaje = sMensaje + mo_AdminServiciosHosp.MensajeError
       
       If sMensaje <> "" Then
           MsgBox sMensaje, vbInformation, Me.Caption
       End If
    
       mo_cmbIdTipoServicio.BoundText = ml_idTipoServicio
        
       Set cmdCodigoHIS.ListSource = mo_ReglasComunes.UPServiciosSeleccionarTodos

       mo_cmbIdTipoEdad.BoundColumn = "IdTipoEdad"
       mo_cmbIdTipoEdad.ListField = "DescripcionLarga"
       Set mo_cmbIdTipoEdad.RowSource = mo_ReglasComunes.TiposEdadSeleccionarTodos
       
       If mi_Opcion = sghModificar Then
            Set oRsConsultoriosMenosElactual = mo_Facturacion.ServiciosSeleccionarPorFiltro(" idTipoServicio=1 and idservicio<>" & ml_idServicio, sghPorDescripcion)
            mo_cmbConsultorios.BoundColumn = "idServicio"
            mo_cmbConsultorios.ListField = "nombre"
            Set mo_cmbConsultorios.RowSource = oRsConsultoriosMenosElactual
       End If
       
       Set cmbUPSsusalud.ListSource = mo_ReglasComunes.SuSalud_upsSeleccionarTodos
       Set cmbUPSfua.ListSource = mo_ReglasComunes.SisFuaUPServiciosSeleccionarTodos
End Sub

Private Sub cmbIdDepartamento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub cmbIdDepartamento_LostFocus()
   If cmbIdDepartamento.Text <> "" Then
       mo_cmbIdDepartamento.BoundText = Val(Split(cmbIdDepartamento.Text, " = ")(0))
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
End Sub

Private Sub cmbIdTipoServicio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdDepartamento_Click()
Dim sMensaje As String

       mo_cmbIdEspecialidad.BoundColumn = "IdEspecialidad"
       mo_cmbIdEspecialidad.ListField = "DescripcionLarga"
       Set mo_cmbIdEspecialidad.RowSource = mo_AdminServiciosHosp.EspecialidadesSeleccionarporDepartamento(Val(mo_cmbIdDepartamento.BoundText))
       
       mo_cmbIdEspecialidad.BoundText = ""
       
       If mo_AdminServiciosHosp.MensajeError <> "" Then
        MsgBox mo_AdminServiciosHosp.MensajeError, vbInformation, Me.Caption
       End If

End Sub

Private Sub cmbIdEspecialidad_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdEspecialidad
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdEspecialidad_LostFocus()
   If cmbIdEspecialidad.Text <> "" Then
       mo_cmbIdEspecialidad.BoundText = Val(Split(cmbIdEspecialidad.Text, " = ")(0))
   End If
End Sub

Private Sub cmbIdEspecialidad_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub




Private Sub cmbTipoSexo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbTipoSexo
    AdministrarKeyPreview KeyCode

End Sub







Private Sub cmdCodigoHIS_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmdCodigoHIS
    AdministrarKeyPreview KeyCode

End Sub

Private Sub cmdMuestraActividades_Click()
    If Trim(Me.cmdCodigoHIS.Text) <> "" And (mi_Opcion = sghAgregar Or mi_Opcion = sghModificar) Then
            Dim oAdmisionCEatencSimultanea As New AdmisionCEatencSimultanea
            oAdmisionCEatencSimultanea.FormLlamante = "SERVICIOS"
            oAdmisionCEatencSimultanea.UPS = Me.cmdCodigoHIS.Text
            oAdmisionCEatencSimultanea.Show 1
            Set oAdmisionCEatencSimultanea = Nothing
    Else
            MsgBox "Debe elegir el UPS y solo funciona para AGREGAR Y/O MODIFICAR", vbInformation, Me.Caption
    End If
End Sub

Private Sub Form_Initialize()
    Set mo_cmbIdDepartamento.MiComboBox = cmbIdDepartamento
    Set mo_cmbIdTipoServicio.MiComboBox = cmbIdTipoServicio
    Set mo_cmbIdEspecialidad.MiComboBox = cmbIdEspecialidad
    Set mo_cmbIdTipoEdad.MiComboBox = cmbIdTipoEdad
    Set mo_cmbConsultorios.MiComboBox = cmbConsultorios
End Sub





Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtCodigo
AdministrarKeyPreview KeyCode
End Sub

Private Sub txtCodigo_LostFocus()
    txtCodigo = UCase(Trim(txtCodigo))
   mo_Formulario.MarcarComoVacio txtCodigo
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub








Private Sub txtCodigoSEM_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtCodigoSEM
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtEdadMax_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtEdadMax
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtEdadMax_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub



Private Sub txtEdadMin_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtEdadMin
   AdministrarKeyPreview KeyCode

End Sub



Private Sub txtMaxCitasAdelantadas_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtMaxCitasAdelantadas
    AdministrarKeyPreview KeyCode

End Sub



Private Sub txtMaxCitasAdicionales_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtMaxCitasAdicionales
    AdministrarKeyPreview KeyCode
End Sub



Private Sub txtMaxCitasSISadelantadas_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtMaxCitasSISadelantadas
    AdministrarKeyPreview KeyCode

End Sub


Private Sub txtMaxCitasSisHoy_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtMaxCitasSisHoy
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNombre
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtNombre_LostFocus()
   mo_Formulario.MarcarComoVacio txtNombre
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Servicios
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()

 mo_Formulario.HabilitarDeshabilitar Me.cmbIdTipoServicio, False
 mo_Formulario.HabilitarDeshabilitar Me.txtCodProductoEstancia, False
 mo_Formulario.HabilitarDeshabilitar Me.txtProductoEstancia, False
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

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Servicios
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Servicios"
       Case sghModificar
           Me.Caption = "Modificar Servicios"
           fraConsultorios.Enabled = True
       Case sghConsultar
           Me.Caption = "Consultar Servicios"
       Case sghEliminar
           Me.Caption = "Eliminar Servicios"
       End Select

       Select Case mi_Opcion
       Case sghModificar, sghConsultar, sghEliminar
       End Select
       GenerarRecordsetTemporal
       CargarComboBoxes
       lcBuscaParametro.DevuelveComboLlenoSegunDescripcion cmdFUAanexo, 359
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
       
       chkUsaFormatoFUA_Click
       
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Servicios
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Activate()
   If mi_Opcion <> sghAgregar Then
       If Not mb_ExistenDatos Then
           Me.Visible = False
       End If
   End If
   On Error Resume Next
   'mgaray20140926
   'If mo_cmbIdTipoServicio.BoundText = 1 Then
   If mo_cmbIdTipoServicio.BoundText = sghTipoServicio.sghConsultaExterna Then
      Me.FraConsultorio.Enabled = True
      Me.FraServicios.Enabled = False
   Else
      Me.FraConsultorio.Enabled = False
      Me.FraServicios.Enabled = True
      'mgaray20140926
      mo_Formulario.HabilitarDeshabilitar chkFuaAdmisionEmerg, False
      If mo_cmbIdTipoServicio.BoundText = sghTipoServicio.sghEmergenciaConsultorios Then
            mo_Formulario.HabilitarDeshabilitar chkFuaAdmisionEmerg, True
      End If
   End If
   '
'   If  False Then   'licencia
'      MsgBox lcMensajeLicencia, vbInformation, Me.Caption
'   End If
   '

End Sub
Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub
Private Sub btnAceptar_Click()

If wxFranklin = "*" Then Exit Sub

   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
            CargaDatosAlObjetosDeDatos
           If ValidarReglas() Then
               If AgregarDatos() Then
                    MsgBox "Los datos se agregaron correctamente", vbInformation, Me.Caption
                    LimpiarFormulario
                    Me.cmbIdDepartamento.SetFocus
                Else
                    MsgBox "Hubo error al agregar los datos", vbInformation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
            CargaDatosAlObjetosDeDatos
           If ValidarReglas() Then
               If ModificarDatos() Then
                    MsgBox "Los datos se modificaron correctamente", vbInformation, Me.Caption
                    Me.Visible = False
                Else
                    MsgBox "Hubo error al modificar los datos ", vbInformation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
                CargaDatosAlObjetosDeDatos
               If EliminarDatos() Then
                    MsgBox "Los datos se eliminaron correctamente", vbInformation, Me.Caption
                    Me.Visible = False
                Else
                    MsgBox "Hubo error al eliminar los datos ", vbInformation, Me.Caption
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
       sMensaje = sMensaje + "Ingrese el nombre del servicio" + Chr(13)
   End If
   If Val(mo_cmbIdEspecialidad.BoundText) = 0 Then
       sMensaje = sMensaje + "Ingrese la especialidad" + Chr(13)
   End If
   If Val(mo_cmbIdTipoServicio.BoundText) = 0 Then
       sMensaje = sMensaje + "Ingrese el tipo de servicio" + Chr(13)
   End If
   If Val(mo_cmbIdDepartamento.BoundText) = 0 Then
       sMensaje = sMensaje + "Ingrese el departamento" + Chr(13)
   End If
   If Me.txtCodigo.Text = "" Then
       sMensaje = sMensaje + "Ingrese el código" + Chr(13)
   End If
   If Me.cmbTipoSexo.Text = "" Then
       sMensaje = sMensaje + "Elija el TIPO SEXO" + Chr(13)
   End If
   If Val(Me.txtEdadMax.Text) = 0 Then
       sMensaje = sMensaje + "Ingrese la EDAD MAXIMA" + Chr(13)
   End If
   If Me.cmdCodigoHIS.Text = "" And mo_cmbIdTipoServicio.BoundText = "1" Then
        sMensaje = sMensaje + "Seleccione un item de la lista de UPS (código de Servicio)" + Chr(13)
   End If
   If txtProductoEstancia.Text = "" And mo_cmbIdTipoServicio.BoundText <> "1" Then
        sMensaje = sMensaje + "Elija ESTANCIA POR DIA" + Chr(13)
   End If
   If Me.cmbIdTipoEdad.Text = "" Then
       sMensaje = sMensaje + "Elija  TIPO EDAD" + Chr(13)
   End If
   
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
Dim rsServicios As Recordset

   ValidarReglas = False
   
   Set rsServicios = mo_AdminServiciosHosp.ServiciosObtenerConElMismoCodigo(mo_Servicios)
   If Not (rsServicios.EOF And rsServicios.BOF) Then
        MsgBox "Ya existe un servicio con el mismo código" + Chr(13) + "Servicio: " & rsServicios!nombre + " Especialidad: " + rsServicios!especialidad, vbExclamation, Me.Caption
        rsServicios.Close
        Exit Function
   End If
   
   Set rsServicios = mo_AdminServiciosHosp.ServiciosObtenerConElMismoNombre(mo_Servicios)
   If Not (rsServicios.EOF And rsServicios.BOF) Then
        MsgBox "Ya existe un servicio con el mismo nombre" + Chr(13) + "Código: " & rsServicios!codigo + " Especialidad: " + rsServicios!especialidad, vbExclamation, Me.Caption
        rsServicios.Close
        Exit Function
   End If
   
   ValidarReglas = True
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla Servicios
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()
   Dim oCampos() As String
   Dim lcCodigoHIS As String, lcUPSfua As String, lcUPSsusalud As String
   On Error Resume Next
   '
   If Me.cmdCodigoHIS.Text = "" Then
       lcCodigoHIS = ""
   Else
        oCampos = Split(Me.cmdCodigoHIS.List(cmdCodigoHIS.ListIndex), "|")
        lcCodigoHIS = oCampos(1)
   End If
   '
   If Me.cmbUPSfua.Text = "" Then
       lcUPSfua = ""
   Else
        oCampos = Split(Me.cmbUPSfua.List(cmbUPSfua.ListIndex), "|")
        lcUPSfua = oCampos(1)
   End If
   '
   If Me.cmbUPSsusalud.Text = "" Then
       lcUPSsusalud = ""
   Else
        oCampos = Split(Me.cmbUPSsusalud.List(cmbUPSsusalud.ListIndex), "|")
        lcUPSsusalud = oCampos(1)
   End If
   '
   With mo_Servicios
           .idServicio = Me.idServicio
           .nombre = Me.txtNombre.Text
           .IdEspecialidad = Val(mo_cmbIdEspecialidad.BoundText)
           .idTipoServicio = Val(mo_cmbIdTipoServicio.BoundText)
           .codigo = Me.txtCodigo.Text
           .IdUsuarioAuditoria = Me.idUsuario
           .solotipoSexo = IIf(Me.cmbTipoSexo.ListIndex = 0, 1, IIf(Me.cmbTipoSexo.ListIndex = 1, 2, 3))
           .maximaEdad = DevuelveEdadEnDiasXtipo(Val(Me.txtEdadMax.Text))
           .codigoServicioSem = Me.txtCodigoSEM.Text
           .ubicacionSEM = Me.txtUbicacionSEM.Text
           .codigoServicioHIS = lcCodigoHIS
           .CostoCeroCE = IIf(chkCostoCero.Value = 1, "S", "")
           .minimaEdad = DevuelveEdadEnDiasXtipo(Val(Me.txtEdadMin.Text))
           .IdEstado = IIf(chkEstado.Value = 1, 1, 0)
           .idProducto = Val(Me.txtCodProductoEstancia.Tag)
           .Triaje = IIf(chkTriaje.Value = 1, True, False)
           .EsObservacionEmergencia = IIf(Me.chkEsObsEmerg.Value = 1, True, False)   '09/08/2011
           .UsaModuloMaterno = IIf(Me.chkUsaModuloMaterno.Value = 1, True, False)  '15/04/2013
           .UsaModuloNinoSano = IIf(Me.chkUsaModuloNinoSano.Value = 1, True, False)  '15/04/2013
           'mgaray20140926
            If Val(mo_cmbIdTipoServicio.BoundText) = sghTipoServicio.sghConsultaExterna Then
                .UsaGalenHos = IIf(Me.chkEnGalenHos.Value = 1, True, False)  '16/05/13
            Else
                .UsaGalenHos = IIf(Me.chkFuaAdmisionEmerg.Value = 1, True, False)  '2014/09/22
            End If
           .TipoEdad = Val(mo_cmbIdTipoEdad.BoundText)
           .UsaFUA = IIf(chkUsaFormatoFUA.Value = 1, True, False)
           .codigoServicioFUA = lcUPSfua
           .codigoServicioSuSalud = lcUPSsusalud
           .FuaTipoAnexo2015 = Val(Left(cmdFUAanexo.Text, 1))
           .MaxCuposAdicionales = Val(Me.txtMaxCitasAdicionales.Text)
           .MaxCuposCitasAdelantadas = Val(Me.txtMaxCitasAdelantadas.Text)
           .MaxCuposCitasAdelandatasSIS = Val(Me.txtMaxCitasSISadelantadas.Text)
           .MaxCuposCitasHoySIS = Val(Me.txtMaxCitasSisHoy.Text)
           .usaNacimiento = IIf(chkFichaNacimiento.Value = 1, 1, 0)
           .emergenciaCorrelativo = IIf(.idTipoServicio = 2, txtEmergCorrelativo.Text, "")
           .NoUsaMensajeTexto = IIf(chkNoUsaMTcelular.Value = 1, 1, 0)
   End With
   
End Sub




'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
    AgregarDatos = mo_AdminServiciosHosp.ServiciosAgregar(mo_Servicios, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtNombre.Text)
    ActualizaPuntoCarga
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------
Function ModificarDatos() As Boolean
    ModificarDatos = mo_AdminServiciosHosp.ServiciosModificar(mo_Servicios, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtNombre.Text)
    ActualizaPuntoCarga
    GrabaConsultoriosDeAtencionSimultanea
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean
    GrabaConsultoriosDeAtencionSimultanea
    EliminarDatos = mo_AdminServiciosHosp.ServiciosEliminar(mo_Servicios, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtNombre.Text)
    ActualizaPuntoCarga
    
End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Servicios
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()
       Dim oConexion As New Connection
       Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
       Dim lnFor As Integer
       oConexion.CommandTimeout = 300
       oConexion.CursorLocation = adUseClient
       oConexion.Open sighentidades.CadenaConexion
       
       
       Set mo_Servicios = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(Me.idServicio, oConexion)
        If mo_AdminServiciosHosp.MensajeError <> "" Then
            MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminServiciosComunes.MensajeError, vbInformation, Me.Caption"
            mb_ExistenDatos = False
            Exit Sub
        End If
        
       If Not mo_Servicios Is Nothing Then
           With mo_Servicios
                Me.idServicio = .idServicio
                Me.txtNombre.Text = .nombre
                mo_cmbIdTipoServicio.BoundText = .idTipoServicio
                Me.txtCodigo = .codigo
                Dim doEspecialidad As New DOEspecialidades
                Set doEspecialidad = mo_AdminServiciosHosp.EspecialidadesSeleccionarPorId(.IdEspecialidad)
                mo_cmbIdDepartamento.BoundText = doEspecialidad.IdDepartamento
                
                mo_cmbIdEspecialidad.BoundText = .IdEspecialidad
                Me.cmbTipoSexo.ListIndex = IIf(.solotipoSexo = 1, 0, IIf(.solotipoSexo = 2, 1, 2))
                Me.txtEdadMax.Text = .maximaEdad
                Me.txtCodigoSEM.Text = .codigoServicioSem
                Me.txtUbicacionSEM.Text = .ubicacionSEM
                Me.cmdCodigoHIS.Text = .codigoServicioHIS
                Me.chkCostoCero.Value = IIf(.CostoCeroCE = "S", 1, 0)
                Me.txtEdadMin.Text = .minimaEdad
                Me.txtCodProductoEstancia.Tag = .idProducto
                chkEstado.Value = .IdEstado
                Me.chkTriaje.Value = IIf(.Triaje = True, 1, 0)
                Me.chkEsObsEmerg.Value = IIf(.EsObservacionEmergencia = True, 1, 0)  '09/08/2011
                Me.chkUsaModuloMaterno.Value = IIf(.UsaModuloMaterno = True, 1, 0)  '15/04/2013
                Me.chkUsaModuloNinoSano.Value = IIf(.UsaModuloNinoSano = True, 1, 0)   '15/04/2013
                'mgaray20140926
                If .idTipoServicio = sghTipoServicio.sghConsultaExterna Then
                    Me.chkEnGalenHos.Value = IIf(.UsaGalenHos = True, 1, 0) '16/05/2013
                Else
                    Me.chkFuaAdmisionEmerg.Value = IIf(.UsaGalenHos = True, 1, 0) '2014/09/22
                End If
                mo_cmbIdTipoEdad.BoundText = .TipoEdad
                Me.chkUsaFormatoFUA.Value = IIf(.UsaFUA = True, 1, 0)
                cmbUPSfua.Text = .codigoServicioFUA
                cmbUPSsusalud.Text = .codigoServicioSuSalud
                Me.txtMaxCitasAdelantadas.Text = .MaxCuposCitasAdelantadas
                Me.txtMaxCitasAdicionales.Text = .MaxCuposAdicionales
                Me.txtMaxCitasSISadelantadas.Text = .MaxCuposCitasAdelandatasSIS
                Me.txtMaxCitasSisHoy.Text = .MaxCuposCitasHoySIS
                chkFichaNacimiento.Value = IIf(.usaNacimiento = 1, 1, 0)
                txtEmergCorrelativo.Text = .emergenciaCorrelativo
                chkNoUsaMTcelular.Value = .NoUsaMensajeTexto
                '
                For lnFor = 0 To cmdFUAanexo.ListCount - 1
                    If Val(Left(cmdFUAanexo.List(lnFor), 1)) = .FuaTipoAnexo2015 Then
                       cmdFUAanexo.ListIndex = lnFor
                       Exit For
                    End If
                Next
                '
                mb_ExistenDatos = True
           End With
           DevuelveEdadSegunTipoGrabado mo_Servicios.minimaEdad, mo_Servicios.maximaEdad, mo_Servicios.TipoEdad
           If mo_ReglasArchivoClinico.ServicioEsPuntoCarga(mo_Servicios.idServicio) = True Then
              Me.chkPuntoCarga.Value = 1
           End If
           If mo_Servicios.idProducto > 0 Then
              Call ObtenerNombreServicio(mo_Servicios.idProducto, Me.txtCodProductoEstancia, Me.txtProductoEstancia)
           End If
           CargaConsultoriosDeAtencionSimultanea oConexion
           
       Else
           mb_ExistenDatos = False
           Exit Sub
       End If
       oConexion.Close
       Set oConexion = Nothing
       Set mo_ReglasArchivoClinico = Nothing
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Servicios
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()
           chkNoUsaMTcelular.Value = 0
           Me.idServicio = 0
           Me.txtNombre.Text = ""
           Me.txtCodigo = ""
           mo_cmbIdEspecialidad.BoundText = ""
           mo_cmbIdDepartamento.BoundText = ""
           Me.cmbTipoSexo.Text = ""
           Me.txtCodigoSEM.Text = ""
           Me.txtEdadMax.Text = "150"
           Me.txtUbicacionSEM.Text = ""
           Me.chkPuntoCarga.Value = 0
           Me.chkCostoCero.Value = 0
           Me.txtEdadMin.Text = "0"
           chkEstado.Value = 1
           chkEsObsEmerg.Value = 0
           Me.chkUsaModuloMaterno.Value = 0
           Me.chkUsaModuloNinoSano.Value = 0
           Me.chkEnGalenHos.Value = 0
           cmbUPSfua.Text = ""
           cmbUPSsusalud.Text = ""
           'mgaray20140926
           Me.chkFuaAdmisionEmerg.Value = 0
           Me.txtMaxCitasAdelantadas.Text = 100
           Me.txtMaxCitasAdicionales.Text = 100
           txtEmergCorrelativo.Text = ""
End Sub



Private Sub txtUbicacionSEM_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtUbicacionSEM
    AdministrarKeyPreview KeyCode

End Sub

Sub ActualizaPuntoCarga()
    mo_ReglasComunes.FactPuntosCargaActualiza1 mo_Servicios.idServicio, mo_Servicios.idTipoServicio, _
                     Me.chkPuntoCarga.Value, _
                     Trim(mo_Servicios.nombre) + IIf(mo_Servicios.idTipoServicio = 1, " (CE)", IIf(mo_Servicios.idTipoServicio = 3, " (H)", " (E)")), _
                     mi_Opcion

    Exit Sub
errPtoCarga:
    MsgBox Err.Description
End Sub

Sub DevuelveEdadSegunTipoGrabado(lnEdadMinima As Long, lnEdadMaxima As Long, lnTipoEdad As Long)
    Select Case lnTipoEdad
    Case sghTipoEdades.sghAño
       Me.txtEdadMax = Round(lnEdadMaxima / 365, 0)
       Me.txtEdadMin = Round(lnEdadMinima / 365, 0)
    Case sghTipoEdades.sghMeses
       Me.txtEdadMax = Round(lnEdadMaxima / 30, 0)
       Me.txtEdadMin = Round(lnEdadMinima / 30, 0)
    Case sghTipoEdades.sghDias
    Case Else
    End Select
End Sub

Function DevuelveEdadEnDiasXtipo(lnEdad As Long) As Long
    Select Case Val(mo_cmbIdTipoEdad.BoundText)
    Case sghTipoEdades.sghAño
       DevuelveEdadEnDiasXtipo = lnEdad * 365
    Case sghTipoEdades.sghMeses
       DevuelveEdadEnDiasXtipo = lnEdad * 30
    Case sghTipoEdades.sghDias
       DevuelveEdadEnDiasXtipo = lnEdad
    Case Else
       DevuelveEdadEnDiasXtipo = 0
    End Select
End Function

Sub GenerarRecordsetTemporal()
    With mRs_ConsultoriosAtencSimultanea
          .Fields.Append "IdServicio", adInteger, 4, adFldIsNullable
          .Fields.Append "Servicio", adVarChar, 100, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
    Set grdConsultorios.DataSource = mRs_ConsultoriosAtencSimultanea
    mo_Apariencia.ConfigurarFilasBiColores Me.grdConsultorios, sighentidades.GrillaConFilasBicolor
    grdConsultorios.Caption = ""
End Sub
Private Sub btnAgregar_Click()
    If Val(mo_cmbConsultorios.BoundText) = 0 Then
       MsgBox "Elija un CONSULTORIO", vbInformation, Me.Caption
       Exit Sub
    End If
    If mRs_ConsultoriosAtencSimultanea.RecordCount > 0 Then
       mRs_ConsultoriosAtencSimultanea.MoveFirst
       mRs_ConsultoriosAtencSimultanea.Find "idServicio=" & mo_cmbConsultorios.BoundText
       If Not mRs_ConsultoriosAtencSimultanea.EOF Then
          MsgBox "Ese CONSULTORIO elegido ya está registrado", vbInformation, Me.Caption
          Exit Sub
       End If
    End If
    mRs_ConsultoriosAtencSimultanea.AddNew
    mRs_ConsultoriosAtencSimultanea.Fields!idServicio = Val(mo_cmbConsultorios.BoundText)
    mRs_ConsultoriosAtencSimultanea.Fields!Servicio = Left(cmbConsultorios.Text, 100)
    mRs_ConsultoriosAtencSimultanea.Update
    mRs_ConsultoriosAtencSimultanea.MoveFirst
End Sub
Private Sub grdConsultorios_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdConsultorios.Bands(0).Columns("IdServicio").Hidden = True
    grdConsultorios.Bands(0).Columns("Servicio").Header.Caption = "Consultorio"
    grdConsultorios.Bands(0).Columns("Servicio").Width = 5500
End Sub

Sub CargaConsultoriosDeAtencionSimultanea(oConexion As Connection)
    Dim oRsTmp1 As New Recordset
    Set oRsTmp1 = mo_ReglasAdmision.ServiciosAtenSimultaneaSeleccionarXidServicio(ml_idServicio, oConexion)
    If oRsTmp1.RecordCount > 0 Then
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
            mRs_ConsultoriosAtencSimultanea.AddNew
            mRs_ConsultoriosAtencSimultanea.Fields!idServicio = oRsTmp1!idServicioAtencionSimultanea
            mRs_ConsultoriosAtencSimultanea.Fields!Servicio = oRsTmp1!nombre
            mRs_ConsultoriosAtencSimultanea.Update
            oRsTmp1.MoveNext
       Loop
       mRs_ConsultoriosAtencSimultanea.MoveFirst
    End If
    oRsTmp1.Close
    Set oRsTmp1 = Nothing
End Sub

Sub GrabaConsultoriosDeAtencionSimultanea()
       On Error GoTo ErrGCDAS
       Dim oConexion As New Connection
       oConexion.CursorLocation = adUseClient
       oConexion.CommandTimeout = 300
       oConexion.Open sighentidades.CadenaConexion
       oConexion.BeginTrans
       If mo_ReglasAdmision.ServiciosAtenSimultaneaEliminar(ml_idServicio, oConexion) Then
       End If
       If mi_Opcion = sghModificar Then
          If mRs_ConsultoriosAtencSimultanea.RecordCount > 0 Then
            mRs_ConsultoriosAtencSimultanea.MoveFirst
            Do While Not mRs_ConsultoriosAtencSimultanea.EOF
               If mo_ReglasAdmision.ServiciosAtenSimultaneaAgregar(ml_idServicio, mRs_ConsultoriosAtencSimultanea.Fields!idServicio, _
                                                                   oConexion) Then
               End If
               mRs_ConsultoriosAtencSimultanea.MoveNext
            Loop
          End If
       End If
       oConexion.CommitTrans
       oConexion.Close
       Set oConexion = Nothing
       Exit Sub
ErrGCDAS:
    oConexion.RollbackTrans
    oConexion.Close
    Set oConexion = Nothing
    MsgBox Err.Description
End Sub



Private Sub cmbUPSfua_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbUPSfua
    AdministrarKeyPreview KeyCode
End Sub






Private Sub cmbUPSsusalud_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbUPSsusalud
    AdministrarKeyPreview KeyCode
End Sub
