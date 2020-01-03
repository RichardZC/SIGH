VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{22ACD161-99EB-11D2-9BB3-00400561D975}#1.0#0"; "PVCALE~1.OCX"
Object = "{8FFC5771-EE23-11D3-9DC0-00A0CC3A1AD6}#1.0#0"; "PVDAYV~1.OCX"
Begin VB.UserControl ucCitasLista 
   ClientHeight    =   8445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15345
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   8445
   ScaleWidth      =   15345
   Begin VB.CommandButton cmdCitaAdicional 
      Caption         =   "Cita Adicional <<ca>>"
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
      Left            =   210
      TabIndex        =   26
      Top             =   7890
      Width           =   1965
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   60
      TabIndex        =   24
      Top             =   600
      Width           =   3885
      Begin VB.CommandButton cmdBuscarEspecialidad 
         Caption         =   "Busc. Esp"
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
         TabIndex        =   32
         Top             =   960
         Width           =   1185
      End
      Begin VB.CommandButton cmdCupoMasProximo 
         Caption         =   "Cupo Proximo"
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
         Left            =   2445
         TabIndex        =   33
         Top             =   960
         Width           =   1305
      End
      Begin VB.CommandButton cmdLimpiar 
         Caption         =   "Limpiar"
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
         TabIndex        =   31
         Top             =   960
         Width           =   945
      End
      Begin VB.ComboBox cmbIdServicio 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   810
         TabIndex        =   25
         Top             =   600
         Width           =   3000
      End
      Begin PVCOMBOLibCtl.PVComboBox cmbMedicos 
         Height          =   300
         Left            =   810
         TabIndex        =   30
         Top             =   210
         Width           =   3000
         _Version        =   524288
         _cx             =   5292
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
            Size            =   8.25
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
         Column0.Heading =   "Id"
         Column0.Width   =   20
         Column0.Alignment=   0
         Column0.Hidden  =   -1  'True
         Column0.Name    =   "Idmedico"
         Column0.Format  =   ""
         Column0.Bound   =   -1  'True
         Column0.Locked  =   0   'False
         Column0.HeaderAlignment=   0
         Column1.Heading =   "Médico"
         Column1.Width   =   200
         Column1.Alignment=   0
         Column1.Hidden  =   0   'False
         Column1.Name    =   "Nombre"
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   150
         TabIndex        =   29
         Top             =   270
         Width           =   570
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Consult"
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
         TabIndex        =   28
         Top             =   630
         Width           =   600
      End
   End
   Begin VB.CommandButton btnRefrescar 
      Caption         =   "Refrescar"
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
      Left            =   2430
      TabIndex        =   18
      Top             =   7890
      Width           =   1365
   End
   Begin VB.Frame fraMedico 
      Caption         =   "Médicos programados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6405
      Left            =   60
      TabIndex        =   12
      Top             =   1965
      Width           =   3855
      Begin VB.ComboBox cmbIdEspecialidadMedico 
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
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   5535
         Width           =   3615
      End
      Begin MSDataListLib.DataList lstMedicos 
         Height          =   285
         Left            =   150
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UltraGrid.SSUltraGrid grdMedicos 
         Height          =   5055
         Left            =   60
         TabIndex        =   23
         Top             =   210
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   8916
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
         Caption         =   "grdMedicos"
      End
      Begin VB.Label lblEspecialidadMedico 
         AutoSize        =   -1  'True
         Caption         =   "Especialidad del médico"
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
         TabIndex        =   13
         Top             =   5310
         Width           =   1905
      End
   End
   Begin VB.Frame fraFiltro 
      Caption         =   "Filtro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   3990
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   4155
      Begin VB.CommandButton cmdCerrarBusEspec 
         Caption         =   "Ocultar Busqueda"
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
         TabIndex        =   34
         Top             =   1440
         Width           =   1785
      End
      Begin VB.ComboBox cmbEspecialidad 
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
         Left            =   150
         TabIndex        =   1
         Top             =   1050
         Width           =   3690
      End
      Begin VB.ComboBox cmbDepartamento 
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
         Left            =   165
         TabIndex        =   0
         Top             =   450
         Width           =   3675
      End
      Begin VB.Label Label1 
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
         Left            =   180
         TabIndex        =   10
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label2 
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
         Height          =   345
         Left            =   180
         TabIndex        =   9
         Top             =   810
         Width           =   2235
      End
   End
   Begin VB.Frame fraProgramacion 
      Height          =   7845
      Left            =   3930
      TabIndex        =   7
      Top             =   570
      Width           =   11370
      Begin VB.Frame FraMedicoSeleccionado 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1515
         Left            =   5520
         TabIndex        =   14
         Top             =   3840
         Width           =   5715
         Begin VB.TextBox txtNcuenta 
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
            Left            =   4545
            TabIndex        =   37
            Top             =   195
            Width           =   1095
         End
         Begin VB.CheckBox chkCuposDispMedico 
            Caption         =   "Mostrar CUPOS LIBRES por Médico"
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
            Left            =   120
            TabIndex        =   27
            Top             =   1140
            Width           =   3615
         End
         Begin VB.TextBox txtCuposLibres 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3390
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   570
            Width           =   315
         End
         Begin VB.TextBox txtCuposAsignados 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   570
            Width           =   315
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "N° Cuenta"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   3885
            TabIndex        =   36
            Top             =   240
            Width           =   645
         End
         Begin VB.Label lblCuposSIS 
            AutoSize        =   -1  'True
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   225
            Left            =   135
            TabIndex        =   35
            Top             =   915
            Width           =   135
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cup.Libres"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2550
            TabIndex        =   20
            Top             =   600
            Width           =   675
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Cup.Asignados"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   480
            TabIndex        =   19
            Top             =   600
            Width           =   945
         End
         Begin VB.Image MedicoProg 
            Height          =   330
            Left            =   60
            Picture         =   "ucCitasLista.ctx":0000
            Stretch         =   -1  'True
            Top             =   195
            Width           =   330
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FF6600&
            Height          =   225
            Left            =   3360
            TabIndex        =   17
            Top             =   195
            Width           =   330
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Fecha selec"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2520
            TabIndex        =   16
            Top             =   225
            Width           =   765
         End
         Begin VB.Label lblDiasProg 
            AutoSize        =   -1  'True
            Caption         =   "Días programados al Médico"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   450
            TabIndex        =   15
            Top             =   225
            Width           =   1770
         End
      End
      Begin UltraGrid.SSUltraGrid grdPacientes 
         Height          =   2220
         Left            =   5505
         TabIndex        =   6
         Top             =   5460
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   3916
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   71303188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionAppearance=   "ucCitasLista.ctx":2316
         Caption         =   "Relacion De Pacientes"
      End
      Begin PVDayView.PVDayView Diario 
         Height          =   7425
         Left            =   60
         TabIndex        =   4
         ToolTipText     =   "Haga click con el botón derecho del mouse para agregar una cita o presione la tecla F2"
         Top             =   240
         Width           =   5445
         _Version        =   65536
         DOYAlignment    =   1
         UseCustomCaption=   -1  'True
         Caption         =   ""
         Appearance      =   1
         BorderStyle     =   1
         Increments      =   4
         SelectMode      =   1
         EnableDayChange =   0   'False
         UseControlPanelSettings=   -1  'True
         TimeSeparator   =   ":"
         AMString        =   "a.m."
         PMString        =   "p.m."
         BusinessHoursBegin=   0
         BusinessHoursEnd=   0.5
         TopIndex        =   0
         TimeBackColor   =   16577517
         SelectedTimeBackColor=   8388608
         AppointmentsForeColor=   0
         AppointmentsBackColor=   16777215
         AppointmentsBarColor=   16711680
         BeginProperty TimeFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty AppointmentsFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseStandardDialogs=   0   'False
         FreeTimeColor   =   16777215
         BusyTimeColor   =   16711680
      End
      Begin PVATLCALENDARLib.PVCalendar Calendario 
         Height          =   3645
         Left            =   5520
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Haga click para seleccionar el día que desea asignar la cita"
         Top             =   240
         Width           =   5715
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
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Asignación de citas"
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
      Height          =   495
      Left            =   15
      TabIndex        =   11
      Top             =   15
      Width           =   15255
   End
   Begin VB.Menu mnuAsignacionCitas 
      Caption         =   "mnuAsignacionCitas"
   End
   Begin VB.Menu mnuDiario 
      Caption         =   "mnuDiario"
      Begin VB.Menu mnuDiarioAgregarCita 
         Caption         =   "Agregar Cita"
      End
      Begin VB.Menu mnuModificarDiarioCita 
         Caption         =   "Modificar Cita"
      End
      Begin VB.Menu mnuDiarioConsultarCita 
         Caption         =   "Consultar Cita"
      End
      Begin VB.Menu mnuDiarioEliminarCita 
         Caption         =   "Eliminar Cita"
      End
   End
End
Attribute VB_Name = "ucCitasLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para listar Citas x Consultorios
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Formulario As New sighEntidades.Formulario
Dim oAdmisionCE As New AdmisionCEDetalle
Dim mo_AdminProgramacionMedica As New SIGHNegocios.ReglasDeProgMedica
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminComun As New SIGHNegocios.ReglasComunes
Dim mo_ReglasDeSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_sighProxies As New SIGHProxies.Procesos

Dim mda_UltimoTimeSlotSeleccionado As Date
Dim mda_HoraInicioCita As Date
Dim mda_HoraFinCita As Date

Dim mo_Apariencia As New sighEntidades.GridInfragistic
Dim mo_cmbEspecialidad As New ListaDespleglable
Dim mo_cmbDepartamento As New ListaDespleglable
Dim mo_cmbIdEspecialidadMedico As New ListaDespleglable
Dim mo_cmbIdServicio As New sighEntidades.ListaDespleglable
Dim ml_idUsuario As Long
Dim mb_RefrescandoDiario As Boolean
Dim ml_ContadorCitasDisponibles As Long
Dim ml_ContadorCitasAsignadas As Long
Dim mo_lcNombrePc As String
Dim lbEsUnaCitaAdicional As Boolean
'WCG_2006
Dim mo_CuentasAtencion As New SIGHNegocios.ReglasFacturacion

Private Const COLOR_CUPO_BLOQUEADO = &H80FFFF   'Amarillo
Private Const COLOR_CUPO_SEPARADO = &H1465FC   'naranja
Private Const COLOR_CUPO_DISPONIBLE = &HC0FFC0  'Verde
Private Const COLOR_CUPO_VENCIDO = &HD18D9C     'Morado
Private Const COLOR_DIA_PROGRAMADO = &HC0FFFF
Private Const COLOR_DIA_NO_PROGRAMADO = &HFFFFFF
Private Const TIEMPO_MAX_ESPERA = 3 '(3 HORAS DE ESPERA COMO MAXIMO)
'WCG_2006
Private Const COLOR_CUPO_PAGADO = &HE8926C  'Azul
Dim lcFormaPago As String
Dim oDoCitaBloqueada As New DOCitaBloqueada
Dim mo_lbCargaTablasUnaVez As Boolean
Dim mo_lbNuevoMovimiento As Boolean
Dim lbYaTieneProgramacionCita As Boolean
Dim lcHoraInicioM As String, lcHoraFinM As String, lcHoraFinUltimoCupo As String, ldUltimaFechaSeleccionada As Date
Dim lcTiempoAtencion As String, lbHabilitaCupoAdicional As Boolean
Dim lnIdProgramacion As Long
Dim lnIdTurno As Integer
Dim lbSeUso_FiltroXmedicos As Boolean
'mgaray
Dim DiarioTieneEnfoque As Boolean
'mgaray201504
Dim bNoPresionarBtnRefrescar As Boolean

Dim lnMaxCuposCitasAdelantadas   As Long, lnCuposCitasAdelantadas As Long   'debb-13/05/2016
Dim lnMaxCuposAdicionales As Long, lnCuposAdicionales As Long               'debb-13/05/2016
Dim ldHoy As Date                                                           'debb-13/05/2016
Dim lnMaxCuposCitasAdelantadasSIS As Long, lnCuposCitasAdelantadasSIS As Long   'debb-25/08/2016
Dim lnMaxCuposCitasHoySIS As Long, lnCuposCitasHoySIS As Long                   'debb-25/08/2016
Dim lbTieneDerechoCitasSIS As Boolean, lcMensajeLicencia As String
Dim mi_nroHistoriaCitadoXmedico As Long
Dim mi_idPacienteCitadoXmedico As Long
Dim oRsConsultoriosAsignados As New Recordset
Dim mi_NOCargaDesdeCitas As Boolean
Dim ml_idFuenteFinanciamientoCitadoXmedico As Long, ml_idFormaPagoCitadoXmedico As Long, ml_txtMedicoRefXMedico As String
Dim ml_cmbIdViasAdmisionXmedico As Long, ml_cmbIdTipoReferenciaOrigenXmedico As Long, ml_txtReferenciaOXmedico As String
Dim ml_txtIdEstablecimientoOrigenXmedico As String, ml_cmbServicioReferenciaOXmedico As String, ml_txtDxReferenciaXmedico As Long
Dim ml_lcCodigoEstablecimientoAdscripcionSISxMedico As String

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



Property Let idFormaPagoCitadoXmedico(lValue As Long)
   ml_idFormaPagoCitadoXmedico = lValue
End Property
Property Let idFuenteFinanciamientoCitadoXmedico(lValue As Long)
   ml_idFuenteFinanciamientoCitadoXmedico = lValue
End Property
Property Let NOCargaDesdeCitas(lValue As Boolean)
   mi_NOCargaDesdeCitas = lValue
End Property
'franklin 2017
Property Let idPacienteCitadoXmedico(lValue As Long)
    mi_idPacienteCitadoXmedico = lValue
End Property

'franklin 2017
Property Let nroHistoriaCitadoXmedico(lValue As Long)
    mi_nroHistoriaCitadoXmedico = lValue
End Property

'franklin 2017
Property Let IdMedicoAtencion(lValue As Long)
    Dim oRs As ADODB.Recordset
    Set oRs = cmbMedicos.ListSource
    Dim lIndiceActual As Long
    lIndiceActual = buscarPosicionMedido(lValue, oRs)
    If cmbMedicos.ListIndex <> lIndiceActual Then
        cmbMedicos.ListIndex = lIndiceActual
    End If
    Set oRs = Nothing
    cmbMedicos_KeyPress 13
    'CargaProgramacionYCitasDelMedico lValue
    
    fraProgramacion.Left = 0
    fraProgramacion.Width = UserControl.Width
    Diario.Width = 2500
    'If Diario.Visible = True Then
        Calendario.Left = Diario.Width + 100
        Calendario.Width = UserControl.Width - Diario.Width - 100
    'Else
     '   Calendario.Left = 0
     '   Calendario.Width = UserControl.Width - 100
    'End If
    grdPacientes.Visible = False
    Frame2.Visible = False
    fraMedico.Visible = False
    UserControl.btnRefrescar.Visible = False
    UserControl.cmdCitaAdicional.Visible = False
    FraMedicoSeleccionado.Visible = False
    grdPacientes.Visible = False
End Property

'franklin 2017
Public Function InhabilitaDiario()
    Diario.Visible = False
End Function
Property Let lbNuevoMovimiento(lValue As Boolean)
   mo_lbNuevoMovimiento = lValue
End Property
Property Let lbCargaTablasUnaVez(lValue As Boolean)
   mo_lbCargaTablasUnaVez = lValue
   oAdmisionCE.lbCargaTablasUnaVez = lValue
End Property

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property

Property Get Departamento() As DataCombo
   Set DiarioVista = UserControl.cmbDepartamento
End Property
Property Get Especialidad() As DataCombo
   Set Especialidad = UserControl.cmbEspecialidad
End Property
Property Get DiarioVista() As PVDayView.PVDayView
   Set DiarioVista = UserControl.Diario
End Property
Property Get CalendarioVista() As PVCalendar
   Set CalendarioVista = UserControl.Calendario
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let MenuAgregarEnabled(bValue As Boolean)
   UserControl.mnuDiarioAgregarCita.Enabled = bValue
End Property
Property Let MenuModificarEnabled(bValue As Boolean)
   UserControl.mnuModificarDiarioCita.Enabled = bValue
End Property
Property Let MenuEliminarEnabled(bValue As Boolean)
   UserControl.mnuDiarioEliminarCita.Enabled = bValue
End Property
Property Let MenuConsultarEnabled(bValue As Boolean)
   UserControl.mnuDiarioConsultarCita.Enabled = bValue
End Property

Private Sub btnRefrescar_Click()
    cmdCitaAdicional.Enabled = False
    RefrescarCitasPorMedico
    'mgaray20141014
'    Set cmbMedicos.ListSource = mo_AdminProgramacionMedica.MedicosFiltrarPorDptoYEspecialidadsql2000(0, 0)

    'mgaray201503
    'Set cmbMedicos.ListSource = mo_AdminProgramacionMedica.MedicosFiltrarPorDptoYEspecialidadConEspecialidad(0, 0)
    Call CargarMedicosActivos(Val(mo_cmbDepartamento.BoundText), Val(mo_cmbEspecialidad.BoundText))
End Sub

Private Sub Calendario_Change(ByVal NewDate As Date)
    ldUltimaFechaSeleccionada = NewDate
    Diario.CurrentDate = NewDate
    Diario.Caption = Format(NewDate, "dddd, MMMM dd, yyyy")
    FraMedicoSeleccionado.Caption = ""
    Set grdPacientes.DataSource = Nothing
    Dim sTemp As String
'    sTemp = lstMedicos.BoundText
    RefrescarListaMedicos
'    lstMedicos.BoundText = sTemp
    Dim oRsTmp As New Recordset
    Set oRsTmp = grdMedicos.DataSource
    If oRsTmp.RecordCount = 0 Then
       lstMedicos_Click
    Else
        If bNoPresionarBtnRefrescar = False Then
            btnRefrescar_Click
        End If
    End If
    Set oRsTmp = Nothing

    If mi_idPacienteCitadoXmedico > 0 Then
       mi_NOCargaDesdeCitas = True
    End If
End Sub

Private Sub chkCuposDispMedico_Click()
    If chkCuposDispMedico.Value = 1 Then
      chkCuposDispMedico.Caption = Left("CUPOS LIBRES DE: " & lstMedicos, 45)
    Else
      chkCuposDispMedico.Caption = "Mostrar CUPOS LIBRES por Médico"
    End If
    grdMedicos_Click
End Sub

Private Sub cmbDepartamento_Click()
       
        mo_cmbEspecialidad.BoundColumn = "IdEspecialidad"
        mo_cmbEspecialidad.ListField = "DescripcionLarga"
        On Error Resume Next
        Set mo_cmbEspecialidad.RowSource = mo_AdminServiciosHosp.EspecialidadesSeleccionarporDepartamento(Val(mo_cmbDepartamento.BoundText))
       
        mo_cmbEspecialidad.BoundText = ""
        'mgaray201503
        Call CargarMedicosActivos(Val(mo_cmbDepartamento.BoundText), 0)
        RefrescarListaMedicos
End Sub

Private Sub cmbDepartamento_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbDepartamento_LostFocus()
   If cmbDepartamento.Text <> "" Then
       mo_cmbDepartamento.BoundText = Val(Split(cmbDepartamento.Text, " = ")(0))
   End If
End Sub

Private Sub cmbEspecialidad_Change()
    If cmbEspecialidad.Text <> "" Then
        mo_cmbEspecialidad.BoundText = Val(Split(cmbEspecialidad.Text, " = ")(0))
    End If
    RefrescarListaMedicos
End Sub

Sub RefrescarListaMedicos()
    Dim lnIdMedicoFiltro As Long
    Dim oCampos() As String
    lnIdMedicoFiltro = 0
    If cmbMedicos.ListIndex >= 0 Then
        If cmbMedicos.List(cmbMedicos.ListIndex) <> "" Then
            oCampos = Split(cmbMedicos.List(cmbMedicos.ListIndex), "|")
            lnIdMedicoFiltro = Val(oCampos(0))
       End If
    End If
    '
    lstMedicos.BoundColumn = "IdMedico"
    lstMedicos.ListField = "Nombre"
    Set lstMedicos.RowSource = mo_AdminProgramacionMedica.MedicosFiltrarPorProgramacionYidTipoServicio(Val(mo_cmbDepartamento.BoundText), Val(mo_cmbEspecialidad.BoundText), 0, Diario.CurrentDate, 1, lnIdMedicoFiltro)
    Set grdMedicos.DataSource = mo_AdminProgramacionMedica.MedicosFiltrarPorProgramacionYidTipoServicio(Val(mo_cmbDepartamento.BoundText), Val(mo_cmbEspecialidad.BoundText), Val(mo_cmbIdServicio.BoundText), Diario.CurrentDate, 1, lnIdMedicoFiltro)
    'solamente Consultorios asignados
    If wxParametro522 = "S" Then
       Dim oRsTmp860 As New Recordset
       Dim oRsTmp861 As New Recordset
       With oRsTmp861
              .Fields.Append "IdMedico", adInteger, 4, adFldIsNullable
              .Fields.Append "Nombre", adVarChar, 200, adFldIsNullable
              .Fields.Append "HoraInicio", adVarChar, 5, adFldIsNullable
              .Fields.Append "HoraFin", adVarChar, 5, adFldIsNullable
              .Fields.Append "DServicio", adVarChar, 200, adFldIsNullable
              .Fields.Append "IdServicio", adInteger, 4, adFldIsNullable
              .Fields.Append "IdTipoServicio", adInteger, 4, adFldIsNullable
              .Fields.Append "IdProgramacion", adInteger, 4, adFldIsNullable
              .Fields.Append "IdTurno", adInteger, 4, adFldIsNullable
              .Fields.Append "MaxCuposCitasAdelantadas", adInteger, 4, adFldIsNullable
              .Fields.Append "MaxCuposAdicionales", adInteger, 4, adFldIsNullable
              .Fields.Append "MaxCuposCitasHoySIS", adInteger, 4, adFldIsNullable
              .Fields.Append "MaxCuposCitasAdelandatasSIS", adInteger, 4, adFldIsNullable
              .LockType = adLockOptimistic
              .Open
       End With
       If oRsConsultoriosAsignados.RecordCount > 0 Then
            Set oRsTmp860 = grdMedicos.DataSource
            If oRsTmp860.RecordCount > 0 Then
               oRsTmp860.MoveFirst
               Do While Not oRsTmp860.EOF
                    oRsConsultoriosAsignados.MoveFirst
                    oRsConsultoriosAsignados.Find "idServicio=" & oRsTmp860!IdServicio
                    If Not oRsConsultoriosAsignados.EOF Then
                        oRsTmp861.AddNew
                        oRsTmp861!idMedico = oRsTmp860!idMedico
                        oRsTmp861!nombre = oRsTmp860!nombre
                        oRsTmp861!HoraInicio = oRsTmp860!HoraInicio
                        oRsTmp861!HoraFin = oRsTmp860!HoraFin
                        oRsTmp861!DServicio = oRsTmp860!DServicio
                        oRsTmp861!IdServicio = oRsTmp860!IdServicio
                        oRsTmp861!idTipoServicio = oRsTmp860!idTipoServicio
                        oRsTmp861!IdProgramacion = oRsTmp860!IdProgramacion
                        oRsTmp861!IdTurno = oRsTmp860!IdTurno
                        oRsTmp861!MaxCuposCitasAdelantadas = oRsTmp860!MaxCuposCitasAdelantadas
                        oRsTmp861!MaxCuposAdicionales = oRsTmp860!MaxCuposAdicionales
                        oRsTmp861!MaxCuposCitasHoySIS = oRsTmp860!MaxCuposCitasHoySIS
                        oRsTmp861!MaxCuposCitasAdelandatasSIS = oRsTmp860!MaxCuposCitasAdelandatasSIS
                        oRsTmp861.Update
                    End If
                    oRsTmp860.MoveNext
               Loop
            End If
       End If
       Set grdMedicos.DataSource = oRsTmp861
       Set oRsTmp860 = Nothing
       Set oRsTmp861 = Nothing
    End If
    '
    If Val(mo_cmbDepartamento.BoundText) = 0 And Val(mo_cmbEspecialidad.BoundText) = 0 And Val(mo_cmbIdServicio.BoundText) = 0 Then
        'mo_cmbIdServicio.BoundColumn = "IdServicio"
        'mo_cmbIdServicio.ListField = "Dservicio"
        'Set mo_cmbIdServicio.RowSource = mo_AdminProgramacionMedica.MedicosFiltrarPorProgramacionsql2000(Val(mo_cmbDepartamento.BoundText), Val(mo_cmbEspecialidad.BoundText), Val(mo_cmbIdServicio.BoundText), Diario.CurrentDate)
        FiltraConsultorios
    End If
    LimpiarCitasDelDiario
    
    'Limpia el calendario
    lblDiasProg = "Días programados al médico: "
    Dim iNroDiasMes As Integer
    Dim i As Integer
    iNroDiasMes = sighEntidades.diasdelmes(Year(Calendario.Value), Month(Calendario.Value))
    For i = 1 To iNroDiasMes
        'Calendario.DATEBackColor(CDate(i & "/" & Month(Calendario.Value) & "/" & Year(Calendario.Value))) = COLOR_DIA_NO_PROGRAMADO
        Calendario.DATEImage(CDate(i & "/" & Month(Calendario.Value) & "/" & Year(Calendario.Value))) = Nothing
    Next i
    grdMedicos_Click
End Sub

'debb-22/06/2016
Sub FiltraConsultorios()
'        If Val(mo_cmbDepartamento.BoundText) = 0 Then
'           Exit Sub
'        End If
        Dim oRsConsultorios As New Recordset
        Dim oRsConsultoriosSinRepetidos As New Recordset
        Dim lbEsNuevo As Boolean
        Set oRsConsultorios = mo_AdminProgramacionMedica.MedicosFiltrarPorProgramacionsql2000(Val(mo_cmbDepartamento.BoundText), Val(mo_cmbEspecialidad.BoundText), Val(mo_cmbIdServicio.BoundText), Diario.CurrentDate)
        oRsConsultorios.Filter = "idServicio<>null"   'debb-21/11/2016
        With oRsConsultoriosSinRepetidos
            .Fields.Append "IdServicio", adInteger, 4, adFldIsNullable
            .Fields.Append "dservicio", adVarChar, 100, adFldIsNullable
            .LockType = adLockOptimistic
            .Open
        End With
        If oRsConsultorios.RecordCount > 0 Then
           oRsConsultorios.MoveFirst
           Do While Not oRsConsultorios.EOF
              lbEsNuevo = True
              If oRsConsultoriosSinRepetidos.RecordCount > 0 Then
                 oRsConsultoriosSinRepetidos.MoveFirst
                 oRsConsultoriosSinRepetidos.Find "idServicio=" & oRsConsultorios!IdServicio
                 If Not oRsConsultoriosSinRepetidos.EOF Then
                    lbEsNuevo = False
                 End If
              End If
              If lbEsNuevo = True Then
                  oRsConsultoriosSinRepetidos.AddNew
                  oRsConsultoriosSinRepetidos.Fields!IdServicio = oRsConsultorios.Fields!IdServicio
                  oRsConsultoriosSinRepetidos.Fields!DServicio = oRsConsultorios.Fields!DServicio
                  oRsConsultoriosSinRepetidos.Update
              End If
              oRsConsultorios.MoveNext
           Loop
        End If
        mo_cmbIdServicio.BoundColumn = "IdServicio"
        mo_cmbIdServicio.ListField = "Dservicio"
        Set mo_cmbIdServicio.RowSource = oRsConsultoriosSinRepetidos
        Set oRsConsultorios = Nothing
        Set oRsConsultoriosSinRepetidos = Nothing
End Sub

Private Sub cmbEspecialidad_Click()
    If cmbEspecialidad.Text <> "" Then
        mo_cmbEspecialidad.BoundText = Val(Split(cmbEspecialidad.Text, " = ")(0))
    End If
    'mgaray201503
    Call CargarMedicosActivos(Val(mo_cmbDepartamento.BoundText), Val(mo_cmbEspecialidad.BoundText))
    RefrescarListaMedicos
End Sub

Private Sub cmbEspecialidad_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbEspecialidad_LostFocus()
   If cmbEspecialidad.Text <> "" Then
       mo_cmbEspecialidad.BoundText = Val(Split(cmbEspecialidad.Text, " = ")(0))
   End If
   RefrescarListaMedicos
End Sub

Private Sub cmbIdEspecialidadMedico_Change()

   If cmbIdEspecialidadMedico.Text <> "" Then
       mo_cmbIdEspecialidadMedico.BoundText = Val(Split(cmbIdEspecialidadMedico.Text, " = ")(0))
   End If
   RefrescarCitasPorMedico
   
End Sub

Sub RefrescarCitasPorMedico()
Dim cupoSeleccionado As PVAppointment
Dim oConexion As New Connection
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighEntidades.CadenaConexion
        
        
        If mo_cmbIdEspecialidadMedico.BoundText = "" Then
            MsgBox "Seleccione la especialidad del medico"
            Exit Sub
        End If
    
        If mb_RefrescandoDiario = True Then
            Exit Sub
        End If
        If lbYaTieneProgramacionCita = True Then
           Exit Sub
        End If
        
        mb_RefrescandoDiario = True
        
        
        Set cupoSeleccionado = Diario.AppointmentSet.GetSelectedAppointment
        
        LimpiarCitasDelDiario
        
        LeerCitasDisponiblesDelMedicoPorDia oConexion
        LeerCitasBloqueadasDelMedicoPorDia oConexion
        LeerCitasAsignadasDelMedicoPorDia oConexion

        If Not cupoSeleccionado Is Nothing Then
            cupoSeleccionado.Selected = True
        End If

        mb_RefrescandoDiario = False
        
        oConexion.Close
        Set oConexion = Nothing
End Sub

Private Sub cmbIdEspecialidadMedico_Click()
   If cmbIdEspecialidadMedico.Text <> "" Then
       mo_cmbIdEspecialidadMedico.BoundText = Val(Split(cmbIdEspecialidadMedico.Text, " = ")(0))
   End If
   RefrescarCitasPorMedico
End Sub

Private Sub cmbIdEspecialidadMedico_LostFocus()
   If cmbIdEspecialidadMedico.Text <> "" Then
       mo_cmbIdEspecialidadMedico.BoundText = Val(Split(cmbIdEspecialidadMedico.Text, " = ")(0))
   End If
   RefrescarCitasPorMedico
End Sub



Private Sub cmbIdServicio_Click()
    cmbIdServicio_KeyPress 13
End Sub

Private Sub cmbIdServicio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmbMedicos.Text = ""
      RefrescarListaMedicos
   End If
End Sub



Private Sub cmbMedicos_Click()
    cmbMedicos_KeyPress 13
End Sub
'franklin 2017
Sub CargaProgramacionYCitasDelMedico(lnIdMedico_filtro As Long)
            lbSeUso_FiltroXmedicos = True
            CargaProgramacionXmedico lnIdMedico_filtro, True
            If IsDate(ldUltimaFechaSeleccionada) = True Then
               If ldUltimaFechaSeleccionada = 0 Then
                  ldUltimaFechaSeleccionada = Date
               End If
               'mgaray201504
               bNoPresionarBtnRefrescar = True
               Calendario_Change ldUltimaFechaSeleccionada
               bNoPresionarBtnRefrescar = False
            End If

End Sub
'franklin 2017
Private Sub cmbMedicos_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        Dim oCampos() As String, lnIdMedico_filtro As Long
        If cmbMedicos.ListIndex >= 0 Then
            cmbIdServicio.Text = ""
            chkCuposDispMedico.Value = 1
            oCampos = Split(cmbMedicos.List(cmbMedicos.ListIndex), "|")
            lnIdMedico_filtro = Val(oCampos(0))
            lstMedicos.BoundText = lnIdMedico_filtro
            CargaProgramacionYCitasDelMedico lnIdMedico_filtro
'            lbSeUso_FiltroXmedicos = True
'            CargaProgramacionXmedico lnIdMedico_filtro, True
            chkCuposDispMedico.Caption = "CUPOS LIBRES DE: " & oCampos(1)
'            If IsDate(ldUltimaFechaSeleccionada) = True Then
'               If ldUltimaFechaSeleccionada = 0 Then
'                  ldUltimaFechaSeleccionada = Date
'               End If
'               'mgaray201504
'               bNoPresionarBtnRefrescar = True
'               Calendario_Change ldUltimaFechaSeleccionada
'               bNoPresionarBtnRefrescar = False
'            End If
            btnRefrescar_Click
            lbSeUso_FiltroXmedicos = False
            cmbMedicos.Text = oCampos(1)
        End If
     End If
End Sub

Private Sub cmdBuscarEspecialidad_Click()
    If fraFiltro.Visible = True Then
        fraFiltro.Visible = False
    Else
        fraFiltro.Visible = True
        cmbDepartamento.SetFocus
    End If
End Sub

Private Sub cmdCerrarBusEspec_Click()
    fraFiltro.Visible = False
End Sub

Private Sub cmdCitaAdicional_Click()
    If lstMedicos.BoundText = "" Then
        MsgBox "Seleccione un medico", vbInformation, "Asignación de citas"
        Exit Sub
    End If
    Dim dHoraIni As Long
    Dim dHoraFin As Long
    Dim programacion As PVAppointment
    Dim oCita As New DOCita
    Dim oRsTmp As New Recordset
    Dim lnIdServicio As Long
    
   
    dHoraIni = mo_AdminProgramacionMedica.ConvertirAMinutos(lcHoraFinM)
    If sighEntidades.EsHora(lcHoraFinUltimoCupo) = True Then
       dHoraIni = mo_AdminProgramacionMedica.ConvertirAMinutos(lcHoraFinUltimoCupo)
    End If
    dHoraFin = dHoraIni + Val(lcTiempoAtencion)
    If mo_AdminProgramacionMedica.ConvertirAHora(dHoraFin) > "23:59" Then
        MsgBox "El CUPO ADICIONAL no puede exceder de la MEDIANOCHE", vbInformation, "Asignación de citas"
        Set oRsTmp = Nothing
        Exit Sub
    End If
    '
    oCita.HoraInicio = mo_AdminProgramacionMedica.ConvertirAHora(dHoraIni)
    oCita.HoraFin = mo_AdminProgramacionMedica.ConvertirAHora(dHoraFin)
    oCita.IdProgramacion = lnIdProgramacion
    '
    'chequear que el MEDICO no esté programado en otro turno como hora de inicio -DEBB-17/03/2014
    Set oRsTmp = mo_AdminProgramacionMedica.ProgramacionMedicaSeleccionarXFechaMedico(CDate(Format(Diario.CurrentDate, "dd/mm/yyyy")), Val(lstMedicos.BoundText))
    'oRsTmp.Filter = "idProgramacion=" & lnIdProgramacion
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          If oRsTmp.Fields!IdProgramacion = lnIdProgramacion Then
             lnIdServicio = Trim(Str(oRsTmp.Fields!IdServicio))
          ElseIf oCita.HoraInicio >= oRsTmp.Fields!HoraInicio And oCita.HoraInicio <= oRsTmp.Fields!HoraFin Then
             MsgBox "No se puede registrar CUPO ADICIONAL porque el MEDICO ya tiene otro TURNO PROGRAMADO el mismo día", vbInformation, "CITA ADICIONAL"
             oRsTmp.Close
             Set oRsTmp = Nothing
             Exit Sub
          End If
          oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    'chequear que no esté programado en otro turno como hora de inicio -DEBB-17/03/2014
    Set oRsTmp = mo_AdminProgramacionMedica.ProgramacionMedicaSeleccionarXfechaConsultorio(CDate(Format(Diario.CurrentDate, "dd/mm/yyyy")), lnIdServicio)
    oRsTmp.Filter = "idProgramacion<>" & lnIdProgramacion
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          If oCita.HoraInicio >= oRsTmp.Fields!HoraInicio And oCita.HoraInicio <= oRsTmp.Fields!HoraFin Then
             MsgBox "No se puede registrar CUPO ADICIONAL porque el CONSULTORIO ya tiene otro TURNO PROGRAMADO el mismo día", vbInformation, "CITA ADICIONAL"
             oRsTmp.Close
             Set oRsTmp = Nothing
             Exit Sub
          End If
          oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
    '
    
    
    Set programacion = Diario.AppointmentSet.Add(Right("0" & Trim(Str(Diario.AppointmentSet.Count + 1)), 2) & _
                                                 "  CUPO ADICIONAL", _
                                                 CDate(Format(Diario.CurrentDate, "dd/mm/yyyy") + " " + oCita.HoraInicio), _
                                                 CDate(Format(Diario.CurrentDate, "dd/mm/yyyy") + " " + oCita.HoraFin))
    programacion.DataVariant = oCita
    '
    lbEsUnaCitaAdicional = True
    mnuDiarioAgregarCita_Click
    RefrescarListaMedicos
End Sub

Private Sub cmdCupoMasProximo_Click()
    Dim oCampos() As String
    Dim sNombreMedico As String
    Dim lIDMedico As Long, lIdEspecialidad As Long
    
    lIDMedico = 0
    
    If cmbMedicos.ListIndex >= 0 Then
        oCampos = Split(cmbMedicos.List(cmbMedicos.ListIndex), "|")
        lIDMedico = Val(oCampos(0))
        sNombreMedico = oCampos(1)
    End If
    Call BuscarCitaMasProxima(lIDMedico, Val(mo_cmbEspecialidad.BoundText), sNombreMedico, cmbEspecialidad.Text)
End Sub

Private Sub cmdLimpiar_Click()
    cmbMedicos.Text = ""
    cmbIdServicio.Text = ""
    mo_cmbDepartamento.BoundText = ""
    mo_cmbEspecialidad.BoundText = ""
    
    If IsDate(ldUltimaFechaSeleccionada) = True Then
        bNoPresionarBtnRefrescar = True
       Calendario_Change ldUltimaFechaSeleccionada
       bNoPresionarBtnRefrescar = False
    End If
    If ldUltimaFechaSeleccionada = 0 Then
       ldUltimaFechaSeleccionada = Date
    End If
    btnRefrescar_Click
End Sub

Private Sub Diario_AfterDragAppointment(ByVal Appointment As PVDayView.IPVAppointment)
Dim Cita As PVAppointment
Dim lMinutosNuevosIni As Long
Dim lMinutosNuevosFin As Long
Dim lMinutosIni As Long
Dim lMinutosFin As Long
Dim bInterseccion As Boolean
        
        lMinutosNuevosIni = ConvertirAMinutos(Format(Appointment.StartDateTime, sighEntidades.DevuelveHoraSoloFormato_HM))
        lMinutosNuevosFin = ConvertirAMinutos(Format(Appointment.EndDateTime, sighEntidades.DevuelveHoraSoloFormato_HM))
        
        Set Cita = Diario.AppointmentSet.GetFirst()
        Do While Not Cita Is Nothing
        
            If Cita.Key <> Appointment.Key Then
                
                lMinutosIni = ConvertirAMinutos(Format(Cita.StartDateTime, sighEntidades.DevuelveHoraSoloFormato_HM))
                lMinutosFin = ConvertirAMinutos(Format(Cita.EndDateTime, sighEntidades.DevuelveHoraSoloFormato_HM))
                
                'Caso 1
                If (lMinutosNuevosIni <= lMinutosIni) And (lMinutosNuevosFin > lMinutosIni) And (lMinutosNuevosFin <= lMinutosFin) Then
                    bInterseccion = True
                End If
                
                'Caso 2
                If (lMinutosNuevosIni <= lMinutosIni) And (lMinutosNuevosFin >= lMinutosFin) Then
                    bInterseccion = True
                End If
                
                'Caso 3
                If (lMinutosNuevosIni >= lMinutosIni) And (lMinutosNuevosIni < lMinutosFin) And (lMinutosNuevosFin <= lMinutosFin) Then
                    bInterseccion = True
                End If
                
                'Caso 4
                If (lMinutosNuevosIni >= lMinutosIni) And (lMinutosNuevosIni < lMinutosFin) And (lMinutosNuevosFin >= lMinutosFin) Then
                    bInterseccion = True
                End If
                
            End If
            
            Set Cita = Diario.AppointmentSet.GetNext(Cita)
        Loop
    
        If bInterseccion Then
            If MsgBox("Hay un traslape entre dos citas asignadas, desea continuar?", vbQuestion + vbYesNo) = vbYes Then
                ActualizarCita Appointment
            Else
                Appointment.StartDateTime = mda_HoraInicioCita
                Appointment.EndDateTime = mda_HoraFinCita
            End If
        Else
            ActualizarCita Appointment
        End If

End Sub
Sub ActualizarCita(ByVal Appointment As PVDayView.IPVAppointment)
    
    Dim oDOCita As New DOCita
    Set oDOCita = Appointment.DataVariant
    
    oDOCita.HoraInicio = Format(Appointment.StartDateTime, sighEntidades.DevuelveHoraSoloFormato_HM)
    oDOCita.HoraFin = Format(Appointment.EndDateTime, sighEntidades.DevuelveHoraSoloFormato_HM)
    
    If Not mo_AdminAdmision.CitasModificar(oDOCita) Then
        MsgBox "No se pudo realizar el movimiento", vbInformation, "Modificación de cita"
        Appointment.StartDateTime = mda_HoraInicioCita
        Appointment.EndDateTime = mda_HoraFinCita
    End If

End Sub

Function ConvertirAMinutos(sHora As String) As Long
Dim sHoras() As String
        
        sHoras = Split(sHora, ":")
        ConvertirAMinutos = Val(sHoras(0)) * 60 + Val(sHoras(1))
        
End Function

Private Sub Diario_AfterTimeSlotSelectChange(ByVal TimeSlot As Date)
Dim Cita As PVAppointment
        

        Dim programacion As PVAppointment


End Sub


Private Sub Diario_BeforeDragAppointment(ByVal Appointment As PVDayView.IPVAppointment)
    mda_HoraInicioCita = Appointment.StartDateTime
    mda_HoraFinCita = Appointment.EndDateTime
   
End Sub


Private Sub Diario_DblClick()
    If mi_NOCargaDesdeCitas = True Then
       mnuDiarioAgregarCita_Click
       Exit Sub
    End If
    mnuDiarioConsultarCita_Click
End Sub





Private Sub Diario_GotFocus()
'mgaray
    DiarioTieneEnfoque = True
End Sub



Private Sub Diario_KeyPress(ByVal KeyAscii As Integer)
Diario_DblClick
End Sub

Private Sub Diario_LostFocus()
    'mgaray
    DiarioTieneEnfoque = False
End Sub

Private Sub Diario_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    If Button = 2 Then
        PopupMenu mnuDiario
    Else
        Dim programacion As PVAppointment
        Set programacion = Diario.AppointmentSet.GetSelectedAppointment
        If programacion Is Nothing Then
            Exit Sub
        End If
        'mgaray201411f
        If IsEmpty(programacion.DataVariant) Then
            Exit Sub
        End If
        If programacion.DataVariant Is Nothing Then
            Exit Sub
        End If
        If programacion.BackColor = COLOR_CUPO_BLOQUEADO Then
            Exit Sub
        End If
        If programacion.DataVariant.IdCita > 0 Then
            Exit Sub
        End If
        lbYaTieneProgramacionCita = False
        'mgaray
        If DiarioTieneEnfoque = False Then
            RefrescarCitasPorMedico
            If UserControl.ActiveControl.Name = "Diario" Then
                DiarioTieneEnfoque = True
            End If
        End If
        '
        'RefrescarCitasPorMedico

    End If

End Sub


Private Sub Diario_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   lcFormaPago = ""
End Sub

Private Sub Diario_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
lcFormaPago = ""
End Sub

Private Sub grdMedicos_AfterRowActivate()
    On Error Resume Next
    Dim rsRecordset As ADODB.Recordset
    Set rsRecordset = grdMedicos.DataSource
    If rsRecordset!EsActivo = False Then
        If mi_NOCargaDesdeCitas = False Then
           If Not Err.Number > 0 Then
              MsgBox "Ese Médico no está ACTIVO", vbInformation, ""
           End If
        End If
    Else
        
        datosDelMedico rsRecordset
    End If
End Sub

Sub datosDelMedico(rsRecordset As Recordset)
    lstMedicos.BoundText = Trim(Str(rsRecordset("idMedico")))
    grdPacientes.Caption = "Pacientes para  <" & rsRecordset("dServicio") & ">"
    lcHoraInicioM = rsRecordset("horaInicio")
    lcHoraFinM = rsRecordset("horaFin")
    lcTiempoAtencion = mo_AdminServiciosHosp.EspecialidadCEseleccionarIdServicio(rsRecordset("idServicio"))
    lnIdProgramacion = rsRecordset("idProgramacion")
    lnIdTurno = rsRecordset("idTurno")
    
    lnMaxCuposCitasAdelantadas = rsRecordset("MaxCuposCitasAdelantadas")        'debb-13/05/2016
    lnMaxCuposAdicionales = rsRecordset("MaxCuposAdicionales")                  'debb-13/05/2016
    '
    lnMaxCuposCitasHoySIS = rsRecordset("MaxCuposCitasHoySIS")                       'debb-25/08/2016
    lnMaxCuposCitasAdelantadasSIS = rsRecordset("MaxCuposCitasAdelandatasSIS")          'debb-25/08/2016
    If lbTieneDerechoCitasSIS = False Then
       lnMaxCuposCitasHoySIS = 1000
       lnMaxCuposCitasAdelantadasSIS = 1000
       
    End If
    '
End Sub

'Actualizado 15102014
Private Sub grdMedicos_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Cancel = True
End Sub

Private Sub grdMedicos_Click()
    mi_NOCargaDesdeCitas = False

    lbYaTieneProgramacionCita = False
    cmdCitaAdicional.Enabled = False
    grdMedicos_AfterRowActivate
    lstMedicos_Click
    chkCuposDispMedico.Caption = Left("CUPOS LIBRES DE: " & lstMedicos, 40)
End Sub




Private Sub grdMedicos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    On Error Resume Next
    grdMedicos.Bands(0).Columns("IdMedico").Hidden = True
    grdMedicos.Bands(0).Columns("HoraInicio").Hidden = True
    grdMedicos.Bands(0).Columns("HoraFin").Hidden = True
    grdMedicos.Bands(0).Columns("IdServicio").Hidden = True
    grdMedicos.Bands(0).Columns("IdTipoServicio").Hidden = True
    grdMedicos.Bands(0).Columns("idProgramacion").Hidden = True
    grdMedicos.Bands(0).Columns("idTurno").Hidden = True
    
    grdMedicos.Bands(0).Columns("Nombre").Header.Caption = "Médico"
    grdMedicos.Bands(0).Columns("Nombre").Width = 2000
    grdMedicos.Bands(0).Columns("Nombre").Activation = ssActivationActivateNoEdit
    grdMedicos.Bands(0).Columns("DServicio").Header.Caption = "Consultorio"
    grdMedicos.Bands(0).Columns("DServicio").Width = 1500
    grdMedicos.Bands(0).Columns("DServicio").Activation = ssActivationActivateNoEdit

End Sub



Private Sub grdMedicos_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = 13 Then
       grdMedicos_Click
    End If
End Sub

'Actualizado 15102014
Private Sub grdPacientes_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Cancel = True
End Sub

Private Sub grdPacientes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdPacientes.Bands(0).Columns(0).Width = 600
    grdPacientes.Bands(0).Columns(1).Width = 600
    grdPacientes.Bands(0).Columns(2).Width = 1000
    grdPacientes.Bands(0).Columns(3).Width = 1000
    grdPacientes.Bands(0).Columns(4).Width = 1000
    
End Sub

Private Sub LeerDatosDelMedico()
    
    If lstMedicos.BoundText = "" Then
        'MsgBox "Por favor seleccione un medico", vbInformation, "Programación Médica"
        Exit Sub
    End If
    
    'Obtiene las programaciones del medico del mes correspondiente
    RefrescarCitasPorMedico

    
End Sub
'debb-24/08/2016
Sub LeerCitasAsignadasDelMedicoPorDia(oConexion As Connection)
Dim oCitas As Collection
Dim oCita As DOCita
Dim bPrimeraCita As Boolean
Dim lcUltimaHora99 As String, lnIdProgramacion99 As Long, ldFecha00 As Date
        lnCuposCitasAdelantadas = 0: lnCuposAdicionales = 0         'debb-13/05/2016
        lnCuposCitasAdelantadasSIS = 0: lnCuposCitasHoySIS = 0      'debb-25/08/2016
        
        ml_ContadorCitasAsignadas = 0
        'Obtiene las programaciones del medico del mes correspondiente
        Set oCitas = mo_AdminProgramacionMedica.CitasLeerPorMedicoYFecha(Val(lstMedicos.BoundText), Diario.CurrentDate, oConexion)
        If oCitas.Count > 0 Then
            bPrimeraCita = True
            For Each oCita In oCitas
                If oCita.HoraInicio >= lcHoraInicioM And oCita.HoraInicio <= lcHoraFinM And oCita.IdProgramacion = lnIdProgramacion Then
                    If bPrimeraCita Then
                        Diario.TopIndex = oCita.HoraInicio
                        bPrimeraCita = False
                    End If
                    ml_ContadorCitasAsignadas = ml_ContadorCitasAsignadas + 1
                    ReemplazarCitaComoAsignada oCita
                    lcTiempoAtencion = DateDiff("n", oCita.HoraInicio, oCita.HoraFin)
                    lcUltimaHora99 = oCita.HoraFin
                    lnIdProgramacion99 = oCita.IdProgramacion
                    ldFecha00 = oCita.fecha
                End If
            Next
        End If

        If lcHoraInicioM <> "" And lcHoraFinM <> "" Then
           Set grdPacientes.DataSource = mo_AdminProgramacionMedica.CitasSeleeccionarPacientePorMedicoFechaHoras(Val(lstMedicos.BoundText), Diario.CurrentDate, oConexion, lcHoraInicioM, lcHoraFinM)
        End If
        UserControl.txtCuposAsignados.Text = ml_ContadorCitasAsignadas
        UserControl.txtCuposLibres.Text = ml_ContadorCitasDisponibles - ml_ContadorCitasAsignadas
        
        If Val(UserControl.txtCuposAsignados.Text) > 0 And Val(UserControl.txtCuposLibres.Text) = 0 Then
           cmdCitaAdicional.Enabled = True
        ElseIf Val(UserControl.txtCuposLibres.Text) < 0 And _
                   Format(ldHoy, sighEntidades.DevuelveFechaSoloFormato_DMY) = Format(ldFecha00, sighEntidades.DevuelveFechaSoloFormato_DMY) Then
           Dim oDOProgramacionMedica As New DOProgramacionMedica
           Dim oProgramacionMedica As New ProgramacionMedica
           Set oProgramacionMedica.Conexion = oConexion
           oDOProgramacionMedica.IdProgramacion = lnIdProgramacion99
           If oProgramacionMedica.SeleccionarPorId(oDOProgramacionMedica) = True Then
              If lcUltimaHora99 > oDOProgramacionMedica.HoraFin Then
                    oDOProgramacionMedica.HoraFin = lcUltimaHora99
                    If oProgramacionMedica.Modificar(oDOProgramacionMedica) = True Then
                    End If
              End If
           End If
           Set oDOProgramacionMedica = Nothing
           Set oProgramacionMedica = Nothing
        End If
        'debb-25/08/2016
        Dim lnCuposRestantes As Long
        If CDate(Format(Diario.CurrentDate, sighEntidades.DevuelveFechaSoloFormato_DMY)) > ldHoy Then
           lnCuposRestantes = IIf(Val(txtCuposLibres.Text) < (lnMaxCuposCitasAdelantadasSIS - lnCuposCitasAdelantadasSIS), Val(txtCuposLibres.Text), (lnMaxCuposCitasAdelantadasSIS - lnCuposCitasAdelantadasSIS))
           lblCuposSIS.Caption = "QUEDAN   " & Trim(Str(lnMaxCuposCitasAdelantadasSIS - lnCuposCitasAdelantadasSIS)) & "   CUPOS para Pacientes SIS"
        Else
           lnCuposRestantes = IIf(Val(txtCuposLibres.Text) < (lnMaxCuposCitasHoySIS - lnCuposCitasHoySIS), Val(txtCuposLibres.Text), (lnMaxCuposCitasHoySIS - lnCuposCitasHoySIS))
        End If
        lblCuposSIS.Caption = "QUEDAN   " & Trim(Str(lnCuposRestantes)) & "   CUPOS para Pacientes SIS"
        If lbTieneDerechoCitasSIS = False Then lblCuposSIS.Caption = lcMensajeLicencia
        
End Sub
Sub LeerCitasBloqueadasDelMedicoPorDia(oConexion As Connection)
Dim oCitas As Collection
Dim oCita As DOCitaBloqueada
Dim bPrimeraCita As Boolean

        'Obtiene las programaciones del medico del mes correspondiente
        Set oCitas = mo_AdminProgramacionMedica.CitasLeerBloqueadasPorMedicoYFecha(Val(lstMedicos.BoundText), Diario.CurrentDate, oConexion)
        If oCitas.Count > 0 Then
            bPrimeraCita = True
            For Each oCita In oCitas
                If bPrimeraCita Then
                    Diario.TopIndex = oCita.HoraInicio
                    bPrimeraCita = False
                End If
                ReemplazarCitaComoBloqueada oCita
            Next
        End If

End Sub

'debb-13/05/2016
Sub ReemplazarCitaComoAsignada(oCita As DOCita)
Dim Cita As PVAppointment
Dim lKey As Long
Dim sHoras() As String
Dim dHoraIni As Double
Dim dHoraFin As Double
Dim daFechaIni As Date
Dim sDescripcion As String
Dim sTexto As String
'Dim doPaciente As doPaciente
Dim programacion As PVAppointment
Dim oRsTmp9 As New Recordset
Dim oRsTmp91 As New Recordset
''WCG_2006
'Dim doCuentasAtencion As DOCuentaAtencion
'Dim lCuentaAtencion As Long
        
        
        'ml_ContadorCitasDisponibles = ml_ContadorCitasDisponibles + 1
        Set Cita = Diario.AppointmentSet.GetFirst()
        Do While Not Cita Is Nothing
            If Format(Cita.StartDateTime, sighEntidades.DevuelveHoraSoloFormato_HM) = oCita.HoraInicio And Format(Cita.EndDateTime, sighEntidades.DevuelveHoraSoloFormato_HM) = oCita.HoraFin Then
                
                lKey = Cita.Key
                Diario.AppointmentSet.Remove lKey
            
                'Citas
                sHoras = Split(oCita.HoraInicio, ":")
                dHoraIni = Val(sHoras(0)) + Val(sHoras(1)) / 60
                
                sHoras = Split(oCita.HoraFin, ":")
                dHoraFin = Val(sHoras(0)) + Val(sHoras(1)) / 60
                
                Set oRsTmp91 = mo_ReglasDeSeguridad.AuditoriaFiltrarCitasPorIdAtencion(oCita.idAtencion, sghRegistroCitaCE)
                Set oRsTmp9 = mo_AdminAdmision.AtencionesSeleccionarPorIdAtencion(oCita.idAtencion)
                
                Dim sEstado As String
                Select Case oCita.IdEstadoCita
                Case 2
                    sEstado = "ATENDIDO"
                Case 4
                    sEstado = "PAGADO"
                Case 1
                    sEstado = "SEPARADO"
                    'If DateDiff("h", CDate(oCita.FechaSolicitud & " " & oCita.HoraSolicitud), Now) >= 3 Then
                    If CDate(oCita.fecha & " " & oCita.HoraInicio) < Now Then
                        sEstado = "VENCIDO"
                    End If
                End Select
                
                sDescripcion = ""
                If sDescripcion = "" Then
                   sDescripcion = Left(oRsTmp9.Fields!descripcion, 20)
                End If
                '
                sTexto = Left(Cita.Description, 2) & "__________" + sEstado + "        " + sDescripcion + Chr(13)
                sTexto = sTexto + "    " + oCita.HoraInicio + " - " + oCita.HoraFin + "           (N° Cuenta: " + Str(Trim(oRsTmp9.Fields!idCuentaAtencion)) + ")" + Chr(13)
                sTexto = sTexto + "    HC: " & _
                HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(oRsTmp9.Fields!NroHistoriaClinica)), False) & _
                "  (" & Trim(oRsTmp9.Fields!ApellidoPaterno) + " " + Trim(oRsTmp9.Fields!ApellidoMaterno) + " " + Trim(oRsTmp9.Fields!PrimerNombre) & ")"
                If oCita.EsCitaAdicional = True Then
                   sTexto = sTexto & " <ca>"
                   lnCuposAdicionales = lnCuposAdicionales + 1
                End If
                lnCuposCitasAdelantadas = lnCuposCitasAdelantadas + 1
                If oRsTmp9!IdFormaPago = sghTipoFinanciamiento.sghSis Then
                    lnCuposCitasAdelantadasSIS = lnCuposCitasAdelantadasSIS + 1
                    lnCuposCitasHoySIS = lnCuposCitasHoySIS + 1
                End If
                '
                If oRsTmp91.RecordCount > 0 Then
                   If Not IsNull(oRsTmp91!Usuario) Then
                      If Trim(oRsTmp91!Usuario) <> "" Then
                         sTexto = sTexto & " (us: " & Trim(oRsTmp91!Usuario) & ")"
                      End If
                   End If
                   If Not IsNull(oRsTmp91!nombrePC) Then
                      If Trim(oRsTmp91!nombrePC) <> "" Then
                         sTexto = sTexto & " (pc: " & Trim(oRsTmp91!nombrePC) & ")"
                      End If
                   End If
                End If
                oRsTmp91.Close
                '
                Set programacion = Diario.AppointmentSet.Add(sTexto, oCita.fecha + dHoraIni / 24, oCita.fecha + dHoraFin / 24)
                On Error Resume Next
                programacion.DataVariant = oCita
                '
                Select Case oCita.IdEstadoCita
                Case 4
                    programacion.BackColor = COLOR_CUPO_PAGADO
                Case 1
                    programacion.BackColor = COLOR_CUPO_SEPARADO
                    'Aqui va el codigo para validar el vencimiento de una cita
                    
                    If DateDiff("h", CDate(oCita.FechaSolicitud & " " & oCita.HoraSolicitud), Now) >= TIEMPO_MAX_ESPERA Then
                        programacion.BackColor = COLOR_CUPO_VENCIDO
                    ElseIf oCita.EsCitaAdicional = True Then
                        programacion.BackColor = vbRed 'vbWhite
                    End If
                End Select
                
                programacion.ReadOnly = True
                Exit Sub
            End If
            
            Set Cita = Diario.AppointmentSet.GetNext(Cita)
   

   
        Loop

        
Set oRsTmp9 = Nothing
Set oRsTmp91 = Nothing
End Sub
Sub ReemplazarCitaComoBloqueada(oCita As DOCitaBloqueada)
Dim Cita As PVAppointment
Dim lKey As Long
Dim sHoras() As String
Dim dHoraIni As Double
Dim dHoraFin As Double
Dim daFechaIni As Date
Dim sDescripcion As String
Dim sTexto As String
Dim oDOEmpleado As New dOEmpleado
Dim programacion As PVAppointment
        
        
        Set Cita = Diario.AppointmentSet.GetFirst()
        Do While Not Cita Is Nothing
            If Format(Cita.StartDateTime, sighEntidades.DevuelveHoraSoloFormato_HM) = oCita.HoraInicio And Format(Cita.EndDateTime, sighEntidades.DevuelveHoraSoloFormato_HM) = oCita.HoraFin Then
                
                lKey = Cita.Key
                Diario.AppointmentSet.Remove lKey
            
                'Citas
                sHoras = Split(oCita.HoraInicio, ":")
                dHoraIni = Val(sHoras(0)) + Val(sHoras(1)) / 60
                
                sHoras = Split(oCita.HoraFin, ":")
                dHoraFin = Val(sHoras(0)) + Val(sHoras(1)) / 60
                
                Set oDOEmpleado = mo_AdminComun.EmpleadosSeleccionarPorId(oCita.idUsuario)
                sTexto = "BLOQUEADO" + Chr(13) + oCita.HoraInicio + " - " + oCita.HoraFin + Chr(13) + "Por: " + oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres + Chr(13) + "Fecha: " & oCita.FechaBloqueo & " " + oCita.HoraBloqueo
                sTexto = "BLOQUEADO" + Chr(13) + oCita.HoraInicio + " - " + oCita.HoraFin + "  - Fecha: " & oCita.FechaBloqueo & " " + oCita.HoraBloqueo + Chr(13) + "Por: " + oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
                Set programacion = Diario.AppointmentSet.Add(sTexto, oCita.fecha + dHoraIni / 24, oCita.fecha + dHoraFin / 24)
                
                On Error Resume Next
                programacion.DataVariant = oCita
                programacion.BackColor = COLOR_CUPO_BLOQUEADO
                programacion.ReadOnly = True
                Exit Sub
            End If
            
            Set Cita = Diario.AppointmentSet.GetNext(Cita)
            
        Loop

End Sub

Sub LeerCitasDisponiblesDelMedicoPorDia(oConexion As Connection)
Dim oCitas As Collection
Dim oCita As DOCita
Dim programacion As PVAppointment
Dim sHoras() As String
Dim dHoraIni As Double
Dim dHoraFin As Double
Dim daFechaIni As Date
Dim sDescripcion As String
Dim sTexto As String
Dim doPaciente As doPaciente
Dim bPrimeraCita As Boolean

        If lstMedicos.BoundText = "" Then
            Exit Sub
        End If
        ml_ContadorCitasDisponibles = 0
        'Obtiene las programaciones del medico del mes correspondiente
        'debb-23/04/2018
        Set oCitas = mo_AdminProgramacionMedica.CitasSeleccionarDisponiblesPorMedicoEspecialidadYFechaHoras(Val(lstMedicos.BoundText), Val(mo_cmbIdEspecialidadMedico.BoundText), Diario.CurrentDate, oConexion, lcHoraInicioM, lcHoraFinM, _
                                                                                        oRsConsultoriosAsignados, wxParametro522)
        If oCitas.Count > 0 Then
            lbYaTieneProgramacionCita = True
            bPrimeraCita = True
            For Each oCita In oCitas
                ml_ContadorCitasDisponibles = ml_ContadorCitasDisponibles + 1
                If bPrimeraCita Then
                    Diario.BusinessHoursBegin = Format(oCita.HoraInicio, sighEntidades.DevuelveHoraSoloFormato_HMS)
                    Diario.TopIndex = Format(oCita.HoraInicio, sighEntidades.DevuelveHoraSoloFormato_HMS)
                    bPrimeraCita = False
                End If
                
                sHoras = Split(oCita.HoraInicio, ":")
                dHoraIni = Val(sHoras(0)) + Val(sHoras(1)) / 60
                
                sHoras = Split(oCita.HoraFin, ":")
                dHoraFin = Val(sHoras(0)) + Val(sHoras(1)) / 60
                
                Set programacion = Diario.AppointmentSet.Add(Right("00" & ml_ContadorCitasDisponibles, 2) & "__________ Disponible " + Chr(13) & "    " & oCita.HoraInicio & " - " & oCita.HoraFin, Diario.CurrentDate + dHoraIni / 24, Diario.CurrentDate + dHoraFin / 24)
                programacion.BackColor = COLOR_CUPO_DISPONIBLE
                programacion.ReadOnly = True
                
                sDescripcion = oCita.HoraFin
                lcHoraFinUltimoCupo = oCita.HoraFin
                On Error Resume Next
                programacion.DataVariant = oCita
                DoEvents
            Next
            Diario.BusinessHoursEnd = sDescripcion 'oCita.HoraFin
        Else
            'MsgBox "El médico no tiene programacion para el día", vbInformation, "Asignación de citas"
        End If


End Sub


Private Sub lstMedicos_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

'Private Sub TimerDeCupos_Timer()
'    RefrescarCitasPorMedico
'End Sub

Public Function Inicializar()
    '
    lbTieneDerechoCitasSIS = True
    '
    mo_Formulario.HabilitarDeshabilitar txtCuposAsignados, False
    mo_Formulario.HabilitarDeshabilitar txtCuposLibres, False
    'On Error Resume Next
    'Diario.TopIndex = "07:00:00 a.m."
    Diario.TopIndex = "07:00:00"
    Calendario.AttachDayView Diario
    
    Diario.AttachCalendar Calendario
    
    mb_RefrescandoDiario = False
    
    Set mo_cmbEspecialidad.MiComboBox = cmbEspecialidad
    Set mo_cmbDepartamento.MiComboBox = cmbDepartamento
    Set mo_cmbIdEspecialidadMedico.MiComboBox = cmbIdEspecialidadMedico
    Set mo_cmbIdServicio.MiComboBox = cmbIdServicio
    
    Select Case sighEntidades.TipoActualizacionDeCupos
    Case "Manual"
        btnRefrescar.Visible = True
        
'        TimerDeCupos.Enabled = False
'        TimerDeCupos.Interval = Val(sighEntidades.IntervaloActualizacionCupos) * 1000
    
    Case "Ambos"
        btnRefrescar.Visible = True
        
'        TimerDeCupos.Enabled = True
'        TimerDeCupos.Interval = Val(sighEntidades.IntervaloActualizacionCupos) * 1000
        
    Case "Automatico"
        btnRefrescar.Visible = False
'        TimerDeCupos.Enabled = True
'        TimerDeCupos.Interval = Val(sighEntidades.IntervaloActualizacionCupos) * 1000
    End Select


    ConfigurarAsignacionCitas

    cmbEspecialidad.ListIndex = -1
    mo_Apariencia.ConfigurarFilasBiColores grdPacientes, sighEntidades.GrillaConFilasBicolor
    mo_Apariencia.ConfigurarFilasBiColores grdMedicos, sighEntidades.GrillaConFilasBicolor
    lcHoraInicioM = "": lcHoraFinM = "": lcHoraFinUltimoCupo = ""

'    lstMedicos.BoundColumn = "IdMedico"
'    lstMedicos.ListField = "Nombre"
    'mgaray20141014
'    Set cmbMedicos.ListSource = mo_AdminProgramacionMedica.MedicosFiltrarPorDptoYEspecialidadsql2000(0, 0)

    'mgaray201503
    'Set cmbMedicos.ListSource = mo_AdminProgramacionMedica.MedicosFiltrarPorDptoYEspecialidadConEspecialidad(0, 0)
    Call cargarTodosLosMedicosActivos

    ldHoy = CDate(lcBuscaParametro.RetornaFechaServidorSQL)     'debb-13/05/2016
    
    wxParametro522 = lcBuscaParametro.SeleccionaFilaParametro(522)
    If wxParametro522 = "S" Then
       Set oRsConsultoriosAsignados = mo_ReglasArchivoClinico.ArchiveroServicioFiltrarPorEmpleado(sighEntidades.Usuario)
       oRsConsultoriosAsignados.Filter = "EsConsultorioAsignado=1"
    End If
End Function





Private Sub txtNcuenta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       If Val(txtNcuenta.Text) > 0 Then
          Dim lcSql As String
          Dim oRsTmp1 As New Recordset
          Dim oRsTmp2 As New Recordset
          Dim oReglasFarmacia As New SIGHNegocios.ReglasFarmacia
          Set oRsTmp1 = oReglasFarmacia.AtencionesSelecionarPorCuenta(Val(txtNcuenta.Text))
          If oRsTmp1.RecordCount > 0 Then
             Calendario_Change oRsTmp1!FechaIngreso
             Set oRsTmp2 = grdMedicos.DataSource
             If oRsTmp2.RecordCount > 0 Then
                lcSql = oRsTmp1!HoraIngreso
                oRsTmp2.MoveFirst
                oRsTmp2.Find "idMedico=" & oRsTmp1!IdMedicoIngreso
                grdMedicos_AfterRowActivate
                grdMedicos_Click
'                Set oRsTmp1 = grdPacientes.DataSource
'                If oRsTmp1.RecordCount > 0 Then
'                   oRsTmp1.MoveFirst
'                   oRsTmp1.Find "hi='" & lcSql & "'"
'                   grdPacientes.Refresh
'                End If
             End If
          End If
          oRsTmp1.Close
          Set oRsTmp1 = Nothing
          Set oRsTmp2 = Nothing
          Set oReglasFarmacia = Nothing
       End If
    End If
End Sub

Private Sub UserControl_Resize()
   On Error Resume Next
   lblNombre.Width = UserControl.Width
   
   fraMedico.Height = UserControl.Height - 2800
   
   grdMedicos.Height = fraMedico.Height - 1000
   'lstMedicos.Height = fraMedico.Height - 1000
   
   lblEspecialidadMedico.Top = lstMedicos.Top + grdMedicos.Height + 40
   'lblEspecialidadMedico.Top = lstMedicos.Top + lstMedicos.Height + 40
   
   btnRefrescar.Top = fraMedico.Top + fraMedico.Height + 60
   cmdCitaAdicional.Top = btnRefrescar.Top
   
   cmbIdEspecialidadMedico.Top = lstMedicos.Top + grdMedicos.Height + 250
   'cmbIdEspecialidadMedico.Top = lstMedicos.Top + lstMedicos.Height + 250
   
   fraProgramacion.Height = fraMedico.Height + 2130
   Diario.Height = fraProgramacion.Height - 330
   fraProgramacion.Width = UserControl.Width - fraMedico.Width - 120
   Calendario.Width = fraProgramacion.Width - 5780      '4280
   
   grdPacientes.Width = Calendario.Width
   FraMedicoSeleccionado.Width = Calendario.Width
   grdPacientes.Height = fraProgramacion.Height - Calendario.Height - 1650
   
End Sub
Private Sub UserControl_Terminate()
    Calendario.AttachDayView Nothing
    Diario.AttachCalendar Nothing
End Sub
Public Sub ConfigurarAsignacionCitas()
    
    mo_cmbDepartamento.BoundColumn = "IdDepartamento"
    mo_cmbDepartamento.ListField = "DescripcionLarga"
    Set mo_cmbDepartamento.RowSource = mo_AdminServiciosHosp.DepartamentosSeleccionarTodos
    '
'    mo_cmbIdServicio.BoundColumn = "IdServicio"
'    mo_cmbIdServicio.ListField = "Nombre"
'    Set mo_cmbIdServicio.RowSource = mo_AdminServiciosHosp.ServiciosSeleccionarPorTipoV2(1)
End Sub
Public Sub mnuModificarDiarioCita_Click()
'Dim oAdmisionCE As New AdmisionCEDetalle
Dim programacion As PVAppointment
    
    Set programacion = Diario.AppointmentSet.GetSelectedAppointment

    If programacion Is Nothing Then
        MsgBox "Seleccione la asignación que desea modificar", vbInformation, "Asignacion de citas"
        Exit Sub
    End If

    If programacion.DataVariant Is Nothing Then
        MsgBox "El cupo seleccionado no esta asignado aún", vbInformation, "Asignacion de citas"
        Exit Sub
    End If
    
    If programacion.BackColor = COLOR_CUPO_BLOQUEADO Then
        MsgBox "La cita esta bloqueada no se puede modificar", vbInformation, "Asignacion de citas"
        Exit Sub
    End If
    
    If programacion.DataVariant.IdCita = 0 Then
        MsgBox "El cupo seleccionado no esta asignado aún", vbInformation, "Asignacion de citas"
        Exit Sub
    End If
    'franklin 2017
    If ChequeaSiEsLaMismaCitaAsignadaPorElMedico(programacion.DataVariant.IdCita) = False Then
       MsgBox "El cupo seleccionado no pertenece al Paciente del Médico", vbInformation, "Asignacion de citas"
       Exit Sub
    End If
    If ldHoy > Diario.CurrentDate Then
       MsgBox "No se puede Modificar CITA menor a HOY", vbInformation, "Asignacion de citas"
       Exit Sub
    End If
    
    On Error Resume Next
    oAdmisionCE.LlegoAlMaximoCuposSIS = mo_sighProxies.LLegoAlTopeCitasSIS(Diario.CurrentDate, ldHoy, lnMaxCuposCitasAdelantadasSIS, _
                                        lnCuposCitasAdelantadasSIS, lnMaxCuposCitasHoySIS, lnCuposCitasHoySIS)
    oAdmisionCE.IdCita = programacion.DataVariant.IdCita
    oAdmisionCE.Opcion = sghModificar
    oAdmisionCE.idUsuario = ml_idUsuario
    oAdmisionCE.TipoServicio = sghConsultaExterna
    oAdmisionCE.NroCola = Left(programacion.Description, 2)
    oAdmisionCE.lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroCitaCE
    oAdmisionCE.lcNombrePc = mo_lcNombrePc
    oAdmisionCE.lbNuevoMovimiento = True
    'FRANKLIN 2017
    If mi_nroHistoriaCitadoXmedico > 0 Then
       oAdmisionCE.nroHistoriaCitadoXmedico = 0
       oAdmisionCE.lbCargaTablasUnaVez = True
       oAdmisionCE.idUsuario = sighEntidades.Usuario
    End If
    oAdmisionCE.Show 1
    
    programacion.Description = oAdmisionCE.NombrePaciente
    programacion.DataVariant = oAdmisionCE.Cita
    programacion.BackColor = COLOR_CUPO_SEPARADO
    programacion.ReadOnly = True
    'Unload oAdmisionCE
    
    LeerDatosDelMedico
        
    Diario.AppointmentSet.GetSelectedAppointment.Selected = False
    grdMedicos_Click

End Sub
'debb-25/08/2016
'Function LLegoAlTopeCitasSIS() As String
'    LLegoAlTopeCitasSIS = ""
'    If CDate(Format(Diario.CurrentDate, sighentidades.DevuelveFechaSoloFormato_DMY)) > ldHoy Then
'        If lnMaxCuposCitasAdelantadasSIS <= lnCuposCitasAdelantadasSIS Then
'           LLegoAlTopeCitasSIS = "Ya llegó al tope de CUPOS SIS ADELANTADOS para ese CONSULTORIO (citas mayores a HOY)" & Chr(13) & "que es: " & Trim(Str(lnMaxCuposCitasAdelantadasSIS))
'        End If
'    ElseIf CDate(Format(Diario.CurrentDate, sighentidades.DevuelveFechaSoloFormato_DMY)) = ldHoy Then
'        If lnMaxCuposCitasHoySIS <= lnCuposCitasHoySIS Then
'           LLegoAlTopeCitasSIS = "Ya llegó al tope de CUPOS SIS para ese CONSULTORIO, que es: " & Trim(Str(lnMaxCuposCitasHoySIS))
'        End If
'    End If
'End Function

Public Sub mnuDiarioAgregarCita_Click()
Dim sTexto As String
Dim dHoraIni As Double
Dim dHoraFin As Double
Dim iHoras() As Integer
Dim sHoras() As String
Dim lcDescripcionProg As String
    
    Dim programacion As PVAppointment
    If lbEsUnaCitaAdicional = True Then
       Set programacion = Diario.AppointmentSet.GetLast()
    Else
       Set programacion = Diario.AppointmentSet.GetSelectedAppointment
    End If
    

    If programacion Is Nothing Then
       MsgBox "Seleccione unos de los cupos asignados", vbInformation, "Asignacion de citas"
       lbEsUnaCitaAdicional = False
       Exit Sub
    End If

    If programacion.BackColor = COLOR_CUPO_SEPARADO Then
        MsgBox "El cupo seleccionado ya fue asignado, seleccione un cupo disponible", vbInformation, "Asignacion de citas"
        lbEsUnaCitaAdicional = False
        Exit Sub
    End If
    If programacion.BackColor = COLOR_CUPO_PAGADO Then
        MsgBox "El cupo seleccionado ya fue pagado, seleccione un cupo disponible", vbInformation, "Asignacion de citas"
        lbEsUnaCitaAdicional = False
        Exit Sub
    End If
    If programacion.BackColor = COLOR_CUPO_VENCIDO Then
        MsgBox "El cupo seleccionado ya se ha vencido, debe eliminar esta cita para poder volver a asignar este cupo", vbInformation, "Asignacion de citas"
        lbEsUnaCitaAdicional = False
        Exit Sub
    End If
    If programacion.BackColor = COLOR_CUPO_BLOQUEADO Then
        MsgBox "El cupo está bloqueado, seleccione otro cupo disponible", vbInformation, "Asignacion de citas"
        lbEsUnaCitaAdicional = False
        Exit Sub
    End If
    
    If lstMedicos.BoundText = "" Then
        MsgBox "Seleccione un medico", vbInformation, "Asignación de citas"
        lbEsUnaCitaAdicional = False
        Exit Sub
    End If
    'debb-13/05/2016 (inicio)
    If lbEsUnaCitaAdicional = True Then
        If lnMaxCuposAdicionales <= lnCuposAdicionales Then
           MsgBox "Ya llegó al tope de CITA ADICIONAL <<ca>> para ese CONSULTORIO, que es: " & Trim(Str(lnMaxCuposAdicionales)), vbInformation, ""
           lbEsUnaCitaAdicional = False
           Exit Sub
        End If
    End If
    If CDate(Format(Diario.CurrentDate, sighEntidades.DevuelveFechaSoloFormato_DMY)) > ldHoy Then
        If lnMaxCuposCitasAdelantadas <= lnCuposCitasAdelantadas Then
           MsgBox "Ya llegó al tope de CUPOS ADELANTADOS para ese CONSULTORIO (citas mayores a HOY)" & Chr(13) & "que es: " & Trim(Str(lnMaxCuposCitasAdelantadas)), vbInformation, ""
           lbEsUnaCitaAdicional = False
           Exit Sub
        End If
    End If
    'debb-13/05/2016 (fin)
    If ldHoy > Diario.CurrentDate Then
       MsgBox "No se puede Agregar CITA menor a HOY", vbInformation, "Asignacion de citas"
       Exit Sub
    End If
    
    Dim lcMensajeLicencia As String, lbTieneLicencia As Boolean
    lbTieneLicencia = True
    oAdmisionCE.TieneLicenciaParaMensajeAcelulares = lbTieneLicencia
    
    
    lcDescripcionProg = Left(programacion.Description, 2)
    
    
    oAdmisionCE.LlegoAlMaximoCuposSIS = mo_sighProxies.LLegoAlTopeCitasSIS(Diario.CurrentDate, ldHoy, lnMaxCuposCitasAdelantadasSIS, _
                                        lnCuposCitasAdelantadasSIS, lnMaxCuposCitasHoySIS, lnCuposCitasHoySIS)
    
    oAdmisionCE.FechaIngreso = Format(Diario.CurrentDate, sighEntidades.DevuelveFechaSoloFormato_DMY)
    oAdmisionCE.HoraInicio = Format(programacion.StartDateTime, sighEntidades.DevuelveHoraSoloFormato_HM)
    oAdmisionCE.HoraFin = Format(programacion.EndDateTime, sighEntidades.DevuelveHoraSoloFormato_HM)
    
    oAdmisionCE.idMedico = lstMedicos.BoundText
    oAdmisionCE.NombreMedico = lstMedicos.Text
    oAdmisionCE.TipoServicio = sghConsultaExterna
    oAdmisionCE.IdProgramacion = programacion.DataVariant.IdProgramacion
    
    oAdmisionCE.Opcion = sghAgregar
    oAdmisionCE.idUsuario = ml_idUsuario
    oAdmisionCE.NroCola = lcDescripcionProg
    
    'Bloquear Cita
    oDoCitaBloqueada.fecha = Format(Diario.CurrentDate, sighEntidades.DevuelveFechaSoloFormato_DMY) 'oAdmisionCE.FechaIngreso
    oDoCitaBloqueada.HoraFin = Format(programacion.EndDateTime, sighEntidades.DevuelveHoraSoloFormato_HM) 'oAdmisionCE.HoraFin
    oDoCitaBloqueada.HoraInicio = Format(programacion.StartDateTime, sighEntidades.DevuelveHoraSoloFormato_HM) 'oAdmisionCE.HoraInicio
    oDoCitaBloqueada.idMedico = lstMedicos.BoundText 'oAdmisionCE.idMedico
    oDoCitaBloqueada.idUsuario = ml_idUsuario
    oDoCitaBloqueada.FechaBloqueo = Format(Now, sighEntidades.DevuelveFechaSoloFormato_DMY)
    oDoCitaBloqueada.HoraBloqueo = Format(Now, sighEntidades.DevuelveHoraSoloFormato_HM)
    
    If Not mo_AdminAdmision.CitasBloqueadasAgregar(oDoCitaBloqueada) Then
        MsgBox "No se puede registrar información de bloqueo de citas", vbExclamation, "Asignación de citas "
    End If
    oAdmisionCE.lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroCitaCE
    oAdmisionCE.lcNombrePc = mo_lcNombrePc
    oAdmisionCE.lbNuevoMovimiento = True
    If lbEsUnaCitaAdicional = True Then
       oAdmisionCE.EsCitaAdicional = True
       lbEsUnaCitaAdicional = False
    Else
       oAdmisionCE.EsCitaAdicional = False
    End If
    'FRANKLIN 2017
    If mi_nroHistoriaCitadoXmedico > 0 Then
       oAdmisionCE.idFuenteFinanciamientoCitadoXmedico = ml_idFuenteFinanciamientoCitadoXmedico
       oAdmisionCE.idFormaPagoCitadoXmedico = ml_idFormaPagoCitadoXmedico
       oAdmisionCE.nroHistoriaCitadoXmedico = mi_nroHistoriaCitadoXmedico
       oAdmisionCE.cmbIdViasAdmisionXmedico = ml_cmbIdViasAdmisionXmedico
       oAdmisionCE.cmbIdTipoReferenciaOrigenXmedico = ml_cmbIdTipoReferenciaOrigenXmedico
       oAdmisionCE.txtReferenciaOXmedico = ml_txtReferenciaOXmedico
       oAdmisionCE.txtIdEstablecimientoOrigenXmedico = ml_txtIdEstablecimientoOrigenXmedico
       oAdmisionCE.cmbServicioReferenciaOXmedico = ml_cmbServicioReferenciaOXmedico
       oAdmisionCE.txtDxReferenciaXmedico = ml_txtDxReferenciaXmedico
       oAdmisionCE.txtMedicoRefXMedico = ml_txtMedicoRefXMedico
       oAdmisionCE.lcCodigoEstablecimientoAdscripcionSISxMedico = ml_lcCodigoEstablecimientoAdscripcionSISxMedico
    End If
    
    oAdmisionCE.Show 1
    
    
    If oAdmisionCE.NombrePaciente <> "" Then
         oAdmisionCE.NombrePaciente = oAdmisionCE.NombrePaciente
         sHoras = Split(oAdmisionCE.HoraInicio, ":")
         dHoraIni = CDbl(Val(sHoras(0)) + Val(sHoras(1)) / 60)
         sHoras = Split(oAdmisionCE.HoraFin, ":")
         dHoraFin = CDbl(Val(sHoras(0)) + Val(sHoras(1)) / 60)
         Set programacion = Diario.AppointmentSet.Add(oAdmisionCE.NombrePaciente, Diario.CurrentDate + dHoraIni / 24, Diario.CurrentDate + dHoraFin / 24)
         programacion.DataVariant = oAdmisionCE.Cita
         programacion.BackColor = COLOR_CUPO_SEPARADO
         programacion.ReadOnly = True
    End If
    
    'Desbloquear Cita
    If Not mo_AdminAdmision.CitasBloqueadasEliminar(oDoCitaBloqueada) Then
        MsgBox "No se puede desbloquear citas información de bloqueo de citas" + Chr(13) + mo_AdminAdmision.MensajeError, vbExclamation, "Asignación de citas "
    End If
    
    LeerDatosDelMedico
    On Error Resume Next
    Diario.AppointmentSet.GetSelectedAppointment.Selected = False
    grdMedicos_Click
    TabEnConsultorio
    
    lbEsUnaCitaAdicional = False    'debb-13/05/2016
End Sub

'***************daniel barrantes**************
'***************Busca el NRO. ORDEN para que el PACIENTE se dirija a CAJA y pague
'***************este dato se muestra en la CITA ASIGNADA
Function BuscaNroOrden(lnIdAtencion As Long) As String
    Dim oRs As New ADODB.Recordset
    Dim oConexion As New Connection
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oRs = mo_CuentasAtencion.FactOrdenServicioPagosPorIdAtencion(lnIdAtencion, oConexion)
    BuscaNroOrden = ""
    If oRs.RecordCount > 0 Then
       oRs.Filter = "idPuntoCarga=6"
    End If
    If oRs.RecordCount > 0 Then
       BuscaNroOrden = "(N° Orden Pago: " + Trim(Str(oRs.Fields!IdOrdenPago)) + ")"
    Else
       BuscaNroOrden = ""
    End If
    oRs.Close
    Set oRs = Nothing
    oConexion.Close
    Set oConexion = Nothing
End Function

Function BuscaIdCuentaAtencion(lnIdAtencion As Long) As String
    Dim sSQL As String
    Dim oCn As New ADODB.Connection
    Dim oRs As New ADODB.Recordset
    oCn.Open sighEntidades.CadenaConexion
    Set oRs = mo_AdminAdmision.AtencionesTFinanciamientoXidAtencion(lnIdAtencion, oCn)
    BuscaIdCuentaAtencion = ""
    lcFormaPago = ""
    If oRs.RecordCount > 0 Then
       BuscaIdCuentaAtencion = "Nº Cuenta: " + Trim(Str(oRs.Fields!idCuentaAtencion))
       lcFormaPago = Trim(oRs.Fields!descripcion)
    End If
    oRs.Close
    oCn.Close
End Function



Public Sub mnuDiarioEliminarCita_Click()
'Dim oAdmisionCE As New AdmisionCEDetalle
Dim programacion As PVAppointment
    
    Set programacion = Diario.AppointmentSet.GetSelectedAppointment

    If programacion Is Nothing Then
        MsgBox "Seleccione la asignación que desea modificar", vbInformation, "Asignacion de citas"
        Exit Sub
    End If
    If programacion.DataVariant Is Nothing Then
        MsgBox "El cupo seleccionado no esta asignado aún", vbInformation, "Asignacion de citas"
        Exit Sub
    End If
    If programacion.BackColor = COLOR_CUPO_PAGADO Then
        MsgBox "El cupo seleccionado ya fue pagado no se puede eliminar, seleccione un cupo disponible", vbInformation, "Asignacion de citas"
        Exit Sub
    End If
    If programacion.BackColor = COLOR_CUPO_BLOQUEADO Then
        MsgBox "La cita esta bloqueada no se puede modificar", vbInformation, "Asignacion de citas"
        Exit Sub
    End If
    If programacion.DataVariant.IdCita = 0 Then
        MsgBox "El cupo seleccionado no esta asignado aún", vbInformation, "Asignacion de citas"
        Exit Sub
    End If
    'franklin 2017
    If ChequeaSiEsLaMismaCitaAsignadaPorElMedico(programacion.DataVariant.IdCita) = False Then
       MsgBox "El cupo seleccionado no pertenece al Paciente del Médico", vbInformation, "Asignacion de citas"
       Exit Sub
    End If
    If ldHoy > Diario.CurrentDate Then
       MsgBox "No se puede Eliminar CITA menor a HOY", vbInformation, "Asignacion de citas"
       Exit Sub
    End If
    
    oAdmisionCE.LlegoAlMaximoCuposSIS = ""
    oAdmisionCE.IdCita = programacion.DataVariant.IdCita
    oAdmisionCE.Opcion = sghEliminar
    oAdmisionCE.idUsuario = ml_idUsuario
    oAdmisionCE.TipoServicio = sghConsultaExterna
    oAdmisionCE.NroCola = Val(Left(programacion.Description, 2))
    oAdmisionCE.lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroCitaCE
    oAdmisionCE.lcNombrePc = mo_lcNombrePc
    oAdmisionCE.lbNuevoMovimiento = True
    'FRANKLIN 2017
    If mi_nroHistoriaCitadoXmedico > 0 Then
       oAdmisionCE.lnIdTablaLISTBARITEMS = 102
       oAdmisionCE.nroHistoriaCitadoXmedico = 0
       oAdmisionCE.lbCargaTablasUnaVez = True
       oAdmisionCE.idUsuario = sighEntidades.Usuario
    End If
    oAdmisionCE.Show 1
    
    programacion.BackColor = COLOR_CUPO_DISPONIBLE
    programacion.ReadOnly = True
    
    'Unload oAdmisionCE
    LeerDatosDelMedico

    Diario.AppointmentSet.GetSelectedAppointment.Selected = False
    grdMedicos_Click

End Sub

Public Sub mnuDiarioConsultarCita_Click()
'Dim oAdmisionCE As New AdmisionCEDetalle
Dim programacion As PVAppointment
    
    Set programacion = Diario.AppointmentSet.GetSelectedAppointment

    If programacion Is Nothing Then
        MsgBox "Seleccione la asignación que desea modificar", vbInformation, "Asignacion de citas"
        Exit Sub
    End If
    'mgaray201411f
    If IsEmpty(programacion.DataVariant) Then
        MsgBox "El cupo seleccionado no esta asignado aún", vbInformation, "Asignacion de citas"
        Exit Sub
    End If
    If programacion.DataVariant Is Nothing Then
        MsgBox "El cupo seleccionado no esta asignado aún", vbInformation, "Asignacion de citas"
        Exit Sub
    End If

    If programacion.BackColor = COLOR_CUPO_BLOQUEADO Then
        MsgBox "La cita esta bloqueada no se puede modificar", vbInformation, "Asignacion de citas"
        Exit Sub
    End If

    If programacion.DataVariant.IdCita = 0 Then
        MsgBox "El cupo seleccionado no esta asignado aún", vbInformation, "Asignacion de citas"
        Exit Sub
    End If
    'franklin 2017
    If ChequeaSiEsLaMismaCitaAsignadaPorElMedico(programacion.DataVariant.IdCita) = False Then
       MsgBox "El cupo seleccionado no pertenece al Paciente del Médico", vbInformation, "Asignacion de citas"
       Exit Sub
    End If
    oAdmisionCE.LlegoAlMaximoCuposSIS = ""
    oAdmisionCE.IdCita = programacion.DataVariant.IdCita
    oAdmisionCE.Opcion = sghConsultar
    oAdmisionCE.TipoServicio = sghConsultaExterna
    oAdmisionCE.idUsuario = ml_idUsuario
    oAdmisionCE.NroCola = Left(programacion.Description, 2)
    oAdmisionCE.lbNuevoMovimiento = True
    oAdmisionCE.lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroCitaCE
    'FRANKLIN 2017
    If mi_nroHistoriaCitadoXmedico > 0 Then
       oAdmisionCE.nroHistoriaCitadoXmedico = 0
       oAdmisionCE.lbCargaTablasUnaVez = True
       oAdmisionCE.idUsuario = sighEntidades.Usuario
    End If
    oAdmisionCE.Show 1
    'Unload oAdmisionCE
    'Set oAdmisionCE = Nothing
    Set programacion = Nothing
    Diario.AppointmentSet.GetSelectedAppointment.Selected = False

End Sub

Sub LimpiarCitasDelDiario()
Dim Cita As PVAppointment
Dim lKey As Long
        
        Set Cita = Diario.AppointmentSet.GetFirst()
        Do While Not Cita Is Nothing
            lKey = Cita.Key
            Set Cita = Diario.AppointmentSet.GetNext(Cita)
            Diario.AppointmentSet.Remove lKey
        Loop

End Sub

 Sub TabEnConsultorio()
    On Error Resume Next
    cmbIdServicio.SetFocus
End Sub


'debb-mayo
Private Sub grdMedicos_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
'    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageDown Or KeyCode = vbKeyPageUp Then
'       SePulsoFlechasPaginas
'    End If
End Sub
Private Sub grdMedicos_KeyUp(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageDown Or KeyCode = vbKeyPageUp Then
       SePulsoFlechasPaginas
    End If
End Sub

Sub SePulsoFlechasPaginas()
       On Error Resume Next
       LimpiarCitasDelDiario
       Set grdPacientes.DataSource = Nothing
       grdPacientes.Caption = "Relacion De Pacientes"
       chkCuposDispMedico.Caption = "Mostrar CUPOS LIBRES por Médico"
       grdMedicos_Click
End Sub


Private Sub lstMedicos_Click()
   CargaProgramacionXmedico Val(lstMedicos.BoundText), False
   'actualizado 20140919
   If Not (Diario.AppointmentSet.GetSelectedAppointment Is Nothing) Then
        Diario.AppointmentSet.GetSelectedAppointment.Selected = False
   End If
End Sub

Sub CargaProgramacionXmedico(lnIdMedico_elegido As Long, lbSoloMuestraCuposDisponibles As Boolean)
Dim iNroDiasMes As Integer
Dim i As Integer
Dim oConexion As New Connection
Dim rsDiasProg As New Recordset
Dim oRsTmp1 As New Recordset
Dim lnCuposTotales As Long, lnCuposAsignados As Long, ldFecha As Date, lcCuposLibres As String
Dim lnCuposXdia As Long, lnProgramadosXdia As Long
Dim lbUsoUltimaFechaSeleccionada As Boolean, ldPrimeraFechaProgramada As Date
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion
'    If lbSoloMuestraCuposDisponibles = True Then
'       chkCuposDispMedico.Value = 1
'    End If

    'Limpia el calendario
    lblDiasProg = "Días programados al médico: "
    iNroDiasMes = sighEntidades.diasdelmes(Year(Calendario.Value), Month(Calendario.Value))
    For i = 1 To iNroDiasMes
        Calendario.DATEBackColor(CDate(i & "/" & Month(Calendario.Value) & "/" & Year(Calendario.Value))) = COLOR_DIA_NO_PROGRAMADO
        Calendario.DATEImage(CDate(i & "/" & Month(Calendario.Value) & "/" & Year(Calendario.Value))) = Nothing
        Calendario.DATEText(CDate(i & "/" & Month(Calendario.Value) & "/" & Year(Calendario.Value))) = ""
    Next i

    If lnIdMedico_elegido = 0 Then
        Exit Sub
    End If

    'Selecciona dias programados del mes
'    Dim rsDiasProg As New Recordset
'    Set rsDiasProg = mo_AdminProgramacionMedica.ProgramacionMedicaSeleccionarDiasDeCEPorMedicoYMes(lnIdMedico_elegido, Month(Diario.CurrentDate), Year(Diario.CurrentDate), oConexion)
'    Do While Not rsDiasProg.EOF
'        Calendario.DATEImage(rsDiasProg!fecha) = UserControl.MedicoProg.Picture
'        rsDiasProg.MoveNext
'    Loop
    Set rsDiasProg = mo_AdminProgramacionMedica.ProgramacionMedicaPorIdMedicoMesAnio(lnIdMedico_elegido, Month(Diario.CurrentDate), Year(Diario.CurrentDate), oConexion)
    If rsDiasProg.RecordCount > 0 Then
        Do While Not rsDiasProg.EOF
            If chkCuposDispMedico.Value = 0 Then
                'Calendario.DATEBackColor(rsDiasProg!fecha) = vbGreen
                Calendario.DATEImage(rsDiasProg!fecha) = UserControl.MedicoProg.Picture
            End If
            ldFecha = rsDiasProg!fecha
            lnCuposTotales = 0: lnCuposAsignados = 0: lcCuposLibres = ""
            If lbSeUso_FiltroXmedicos = True Then
               ldPrimeraFechaProgramada = ldFecha
               lbSeUso_FiltroXmedicos = False
            End If
            If ldFecha = ldUltimaFechaSeleccionada And cmbMedicos.Text <> "" Then
               lbUsoUltimaFechaSeleccionada = True
            End If
            Do While Not rsDiasProg.EOF And ldFecha = rsDiasProg!fecha
               If chkCuposDispMedico.Value = 1 And Not IsNull(rsDiasProg!TiempoPromedioAtencion) Then
                    lnCuposXdia = Round(DateDiff("n", rsDiasProg!fecha & " " & rsDiasProg!HoraInicio, rsDiasProg!fecha & " " & rsDiasProg!HoraFin) / rsDiasProg!TiempoPromedioAtencion, 0)
                    lnCuposTotales = lnCuposTotales + lnCuposXdia
                    Set oRsTmp1 = mo_AdminProgramacionMedica.CitasSeleccionarPorIdProgramacion(rsDiasProg!IdProgramacion, oConexion)
                    oRsTmp1.Filter = "idMedico=" & lnIdMedico_elegido
                    lnProgramadosXdia = oRsTmp1.RecordCount
                    lnCuposAsignados = lnCuposAsignados + lnProgramadosXdia
                    If lcCuposLibres = "" Then
                       lcCuposLibres = Left(rsDiasProg!Turno, 2) & ":" & Trim(Str(lnCuposXdia - lnProgramadosXdia))
                    Else
                       lcCuposLibres = lcCuposLibres & " /" & Left(rsDiasProg!Turno, 2) & ":" & Trim(Str(lnCuposXdia - lnProgramadosXdia))
                    End If
               End If
               rsDiasProg.MoveNext
               If rsDiasProg.EOF Then
                  Exit Do
               End If
            Loop
            If chkCuposDispMedico.Value = 1 Then
               'Calendario.DATEText(ldFecha) = "C.Lib= " & Trim(Str(lnCuposTotales - lnCuposAsignados))
               Calendario.DATEText(ldFecha) = lcCuposLibres
            End If
        Loop
        If lbUsoUltimaFechaSeleccionada = False And cmbMedicos.Text <> "" Then
           ldUltimaFechaSeleccionada = ldPrimeraFechaProgramada
        End If
    End If

    'Selecciona especialidades por medico
    mo_cmbIdEspecialidadMedico.BoundColumn = "IdEspecialidad"
    mo_cmbIdEspecialidadMedico.ListField = "DescripcionLarga"
    Dim rsEspecialidad As New Recordset
    Dim lcCabeceraFrame As String
    Set rsEspecialidad = mo_AdminServiciosHosp.EspecialidadesSeleccionarporMedico(lnIdMedico_elegido, oConexion)
    Set mo_cmbIdEspecialidadMedico.RowSource = rsEspecialidad
    
    'FCV10072015
    lcCabeceraFrame = ""
    If Len(lstMedicos.Text) >= 0 Then
        If InStr(lstMedicos.Text, "(") >= 4 Then
            lcCabeceraFrame = Mid(lstMedicos.Text, 1, InStr(lstMedicos.Text, "(") - 3)
        End If
    End If

    If rsEspecialidad.RecordCount = 1 Then
         lbYaTieneProgramacionCita = False
         rsEspecialidad.MoveFirst
         mo_cmbIdEspecialidadMedico.BoundText = rsEspecialidad!IdEspecialidad
'         cmbIdEspecialidadMedico.Enabled = False
         mo_Formulario.HabilitarDeshabilitar cmbIdEspecialidadMedico, False 'FCV10072015
        RefrescarCitasPorMedico
    ElseIf rsEspecialidad.RecordCount > 1 Then
         rsEspecialidad.MoveFirst
         Do While Not rsEspecialidad.EOF
            mo_cmbIdEspecialidadMedico.BoundText = rsEspecialidad!IdEspecialidad
            'cmbIdEspecialidadMedico.Enabled = True
            mo_Formulario.HabilitarDeshabilitar cmbIdEspecialidadMedico, True 'FCV10072015
            If lbYaTieneProgramacionCita = True Then
               Exit Do
            End If
            'Obtiene las programaciones del medico del mes correspondiente
            RefrescarCitasPorMedico
            If ml_ContadorCitasDisponibles > 0 Then
               Exit Do
            End If
            rsEspecialidad.MoveNext
        Loop
    End If
'    lnIdTurno
    lcCabeceraFrame = lcCabeceraFrame & " (Esp:" & cmbIdEspecialidadMedico.Text & ")" & " (Turno:" & DevolverTurno(lnIdTurno) & ")"
    FraMedicoSeleccionado.Caption = lcCabeceraFrame 'FCV10072015

    oConexion.Close
    Set oConexion = Nothing
    Set rsDiasProg = Nothing
    Set oRsTmp1 = Nothing

End Sub
'mgaray201503
Private Function BuscarCitaMasProxima(lIDMedico As Long, lIdEspecialidad As Long, _
                sNombreMedico As String, sNombreEspecialidad As String) As Boolean
    BuscarCitaMasProxima = False
    If lIDMedico = 0 And lIdEspecialidad = 0 Then
        MsgBox "Debe Especificar Médico y/o Especialidad para ubicar una cita", vbInformation, ""
        Exit Function
    End If
    Dim oRs As ADODB.Recordset
    
    Set oRs = mo_AdminProgramacionMedica.ProgramacionMedicaCitaMasProxima(lIDMedico, lIdEspecialidad)
    If Not (oRs Is Nothing) Then
        Dim sMessage As String
        If oRs.RecordCount > 0 Then
            oRs.MoveFirst
            
            sMessage = "Se Encontro Citas Disponibles en la Programación para :"
            If lIDMedico > 0 Then
                sMessage = sMessage & Chr(13) & "Médico : " & sNombreMedico
            End If
            If lIdEspecialidad > 0 Then
                sMessage = sMessage & Chr(13) & "Especialidad : " & sNombreEspecialidad
            End If
            
            sMessage = sMessage & Chr(13) & "Fecha : " & Format(oRs.Fields!fecha, "dd/mm/yyyy")
            
            If MsgBox(sMessage & Chr(13) & "¿Desea visualizar los cupos?", vbInformation + vbYesNo, "") = vbNo Then
                Exit Function
            Else
                UbicarDatosCitaMasProxima oRs
            End If
        Else
            
            sMessage = "No se Encontro Citas Disponibles en la Programacion para :"
            If lIDMedico > 0 Then
                sMessage = sMessage & Chr(13) & "Médico : " & sNombreMedico
            End If
            If lIdEspecialidad > 0 Then
                sMessage = sMessage & Chr(13) & "Especialidad : " & sNombreEspecialidad
            End If
            MsgBox sMessage, vbInformation, ""
            Exit Function
        End If
    End If
    BuscarCitaMasProxima = True
End Function

Private Sub UbicarDatosCitaMasProxima(oRs As ADODB.Recordset)
    Calendario_Change oRs.Fields!fecha
    'Calendario.Value = ss
End Sub

'mgaray201503
Private Sub CargarMedicosActivos(IdDepartamento As Long, IdEspecialidad As Long)
    Dim oCampos() As String, lnIdMedico_filtro As Long
    
    If cmbMedicos.ListIndex >= 0 Then
        oCampos = Split(cmbMedicos.List(cmbMedicos.ListIndex), "|")
        lnIdMedico_filtro = Val(oCampos(0))
    End If
    Dim oRs As ADODB.Recordset, oRsAux As ADODB.Recordset
    
    Set oRs = mo_AdminProgramacionMedica.MedicosFiltrarPorDptoYEspecialidadConEspecialidad( _
                                            IdDepartamento, IdEspecialidad)
                                            
    
                                            
    Set cmbMedicos.ListSource = oRs
    Dim lIndiceActual As Long
    lIndiceActual = buscarPosicionMedido(lnIdMedico_filtro, oRs)
    If cmbMedicos.ListIndex <> lIndiceActual Then
        cmbMedicos.ListIndex = lIndiceActual
    End If
End Sub


'mgaray201504
Private Function buscarPosicionMedido(lIDMedico As Long, oRs As ADODB.Recordset) As Long
    Dim lIndex As Long, lIndiceEncontrado As Long
    
    lIndex = -1
    lIndiceEncontrado = -1
    
    If lIDMedico > 0 Then
        If Not (oRs.EOF = True And oRs.BOF = True) Then
            oRs.MoveFirst
            Do While Not oRs.EOF
                lIndex = lIndex + 1
                If oRs.Fields!idMedico = lIDMedico Then
                    lIndiceEncontrado = lIndex
                End If
                oRs.MoveNext
            Loop
        End If
    End If
    buscarPosicionMedido = lIndiceEncontrado
End Function

Private Sub cargarTodosLosMedicosActivos()
    Call CargarMedicosActivos(0, 0)
End Sub

'FCV 10072015
Public Function DevolverTurno(lnIdTurno As Integer) As String
    Dim mo_Turno As New doTurno
    Set mo_Turno = mo_AdminProgramacionMedica.TurnosSeleccionarPorId(lnIdTurno)
    If mo_AdminProgramacionMedica.MensajeError <> "" Then
        MsgBox "No se pudo obtener los datos" + Chr(13) + mo_AdminProgramacionMedica.MensajeError, vbInformation, "Asignación de citas"
        Exit Function
    End If
    DevolverTurno = ""
    If Not mo_Turno Is Nothing Then
         With mo_Turno
            DevolverTurno = .Codigo & " de " & .HoraInicio & " a " & .HoraFin
         End With
    Else
        Exit Function
    End If
End Function


'franklin 2017
Function ChequeaSiEsLaMismaCitaAsignadaPorElMedico(lnIdCita As Long) As Boolean
    ChequeaSiEsLaMismaCitaAsignadaPorElMedico = False
    If mi_nroHistoriaCitadoXmedico > 0 Then
        Dim oCitas As New Citas
        Dim oDOCita As New DOCita
        Dim oConexion As New Connection
        oConexion.CommandTimeout = 900
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighEntidades.CadenaConexion
        Set oCitas.Conexion = oConexion
        oDOCita.IdCita = lnIdCita
        oDOCita.IdUsuarioAuditoria = sighEntidades.Usuario
        If oCitas.SeleccionarPorId(oDOCita) = True Then
           If oDOCita.idPaciente = mi_idPacienteCitadoXmedico Then
              If CDate(Format(Diario.CurrentDate, sighEntidades.DevuelveFechaSoloFormato_DMY)) > ldHoy Then
                 ChequeaSiEsLaMismaCitaAsignadaPorElMedico = True
              End If
           End If
        Else
        End If
        oConexion.Close
        Set oConexion = Nothing
        Set oDOCita = Nothing
        Set oCitas = Nothing
    Else
        ChequeaSiEsLaMismaCitaAsignadaPorElMedico = True
    End If
End Function

