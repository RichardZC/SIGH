VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{22ACD161-99EB-11D2-9BB3-00400561D975}#1.0#0"; "PVCALE~1.OCX"
Object = "{8FFC5771-EE23-11D3-9DC0-00A0CC3A1AD6}#1.0#0"; "PVDAYV~1.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.UserControl ucSIcitasLista 
   ClientHeight    =   8835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13380
   ScaleHeight     =   8835
   ScaleWidth      =   13380
   Begin VB.Frame fraProgramacion 
      Height          =   8295
      Left            =   3615
      TabIndex        =   10
      Top             =   495
      Width           =   9675
      Begin VB.ListBox lstProgramacion 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   15
         TabIndex        =   47
         Top             =   7260
         Width           =   5805
      End
      Begin PVDayView.PVDayView Diario 
         Height          =   7050
         Left            =   45
         TabIndex        =   11
         ToolTipText     =   "Haga click con el botón derecho del mouse para agregar una programación"
         Top             =   165
         Width           =   5805
         _Version        =   65536
         DOYAlignment    =   2
         UseCustomCaption=   -1  'True
         Caption         =   ""
         Appearance      =   1
         BorderStyle     =   1
         Increments      =   4
         SelectMode      =   1
         EnableDayChange =   0   'False
         UseControlPanelSettings=   0   'False
         TimeSeparator   =   ":"
         AMString        =   "AM"
         PMString        =   "PM"
         BusinessHoursBegin=   0.25
         BusinessHoursEnd=   0.833333333333333
         TopIndex        =   0
         TimeBackColor   =   16577517
         SelectedTimeBackColor=   8388608
         AppointmentsForeColor=   0
         AppointmentsBackColor=   16777215
         AppointmentsBarColor=   16737792
         BeginProperty TimeFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty AppointmentsFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
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
         Height          =   3840
         Left            =   5850
         TabIndex        =   12
         ToolTipText     =   "Seleccione uno o mas días y haga click con el boton derecho de mouse para agregar un programación"
         Top             =   165
         Width           =   3765
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
         DOWText0        =   "Domingo"
         DOWText1        =   "Lunes"
         DOWText2        =   "Martes"
         DOWText3        =   "Miercoles"
         DOWText4        =   "Jueves"
         DOWText5        =   "Viernes"
         DOWText6        =   "Sabado"
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
         DisplayBackColor=   13405544
         DisplayForeColor=   0
         DayBackColor    =   16577517
         DayForeColor    =   0
         SelectedDayForeColor=   16777215
         SelectedDayBackColor=   16737792
         BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
            Name            =   "Tahoma"
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
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UltraGrid.SSUltraGrid grdResumenCpt 
         Height          =   4140
         Left            =   5880
         TabIndex        =   21
         Top             =   4050
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   7303
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Resumen de CPT por día"
      End
      Begin VB.Label Label 
         Caption         =   "Pulsar ENTER para ver pacientes citados"
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
         Height          =   270
         Index           =   0
         Left            =   135
         TabIndex        =   13
         Top             =   150
         Width           =   3930
      End
   End
   Begin VB.Frame fraMedico 
      Height          =   8310
      Left            =   30
      TabIndex        =   1
      Top             =   510
      Width           =   3525
      Begin VB.Frame fraDiasNoLaborables 
         Caption         =   "Días no Laborables por mes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   3255
         TabIndex        =   26
         Top             =   2655
         Visible         =   0   'False
         Width           =   2310
         Begin VB.CommandButton cmdSalir3 
            DisabledPicture =   "ucSiListasCitas.ctx":0000
            DownPicture     =   "ucSiListasCitas.ctx":04C4
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2640
            Picture         =   "ucSiListasCitas.ctx":09B0
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Salir"
            Top             =   2700
            Width           =   660
         End
         Begin VB.CommandButton cmdActualizaDiasNoLaborables 
            DisabledPicture =   "ucSiListasCitas.ctx":0E9C
            DownPicture     =   "ucSiListasCitas.ctx":12FC
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   1875
            Picture         =   "ucSiListasCitas.ctx":1771
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Grabar cambios"
            Top             =   2700
            Width           =   660
         End
         Begin PVATLCALENDARLib.PVCalendar PVCalendarFeriados 
            Height          =   2460
            Left            =   120
            TabIndex        =   42
            ToolTipText     =   "Pulsar CLIC en el día para hacerlo NO LABORABLE, y otro CLIC en el mismo día para hacerlos LABORABLE"
            Top             =   210
            Width           =   3180
            _Version        =   524288
            BorderStyle     =   1
            Appearance      =   1
            FirstDay        =   1
            Frame           =   1
            SelectMode      =   2
            DisplayFormat   =   0
            DateOrientation =   0
            CustomTextOrientation=   2
            ImageOrientation=   0
            DOWText0        =   "Do"
            DOWText1        =   "Lu"
            DOWText2        =   "Ma"
            DOWText3        =   "Mi"
            DOWText4        =   "Ju"
            DOWText5        =   "Vi"
            DOWText6        =   "Sa"
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
            DisplayBackColor=   13405544
            DisplayForeColor=   0
            DayBackColor    =   16577517
            DayForeColor    =   0
            SelectedDayForeColor=   16777215
            SelectedDayBackColor=   16737792
            BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MultiLineText   =   -1  'True
            EditMode        =   0
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblReprogramacion 
            Caption         =   "PUnto Carga"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   405
            Index           =   0
            Left            =   90
            TabIndex        =   39
            Top             =   2685
            Width           =   1680
         End
      End
      Begin VB.Frame fraReprogramacionCita 
         Caption         =   "Reprogramación de Cita"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3405
         Left            =   60
         TabIndex        =   25
         Top             =   1755
         Visible         =   0   'False
         Width           =   3420
         Begin VB.ComboBox cmbSalaNew 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   990
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   2100
            Width           =   2355
         End
         Begin VB.ComboBox txtHoraPac 
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
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   2460
            Width           =   945
         End
         Begin VB.TextBox txtNCita1 
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
            Left            =   2520
            MaxLength       =   99
            TabIndex        =   30
            Top             =   1740
            Width           =   825
         End
         Begin Threed.SSOption optPorFecha 
            Height          =   270
            Left            =   135
            TabIndex        =   43
            Top             =   540
            Width           =   1455
            _ExtentX        =   2566
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
            Caption         =   "Por Fecha"
            Value           =   -1
         End
         Begin VB.CommandButton btnCancelar1 
            DisabledPicture =   "ucSiListasCitas.ctx":1BE6
            DownPicture     =   "ucSiListasCitas.ctx":20AA
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
            Left            =   2700
            Picture         =   "ucSiListasCitas.ctx":2596
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   2805
            Width           =   615
         End
         Begin VB.CommandButton cmdGrabaReprogramacion 
            DisabledPicture =   "ucSiListasCitas.ctx":2A82
            DownPicture     =   "ucSiListasCitas.ctx":2EE2
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
            Left            =   2040
            Picture         =   "ucSiListasCitas.ctx":3357
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   2805
            Width           =   615
         End
         Begin MSMask.MaskEdBox txtFechaCita 
            Height          =   315
            Left            =   1995
            TabIndex        =   27
            Top             =   735
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
         Begin MSMask.MaskEdBox txtFCitaNueva 
            Height          =   315
            Left            =   1995
            TabIndex        =   28
            Top             =   1080
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
         Begin Threed.SSOption optPorNcita 
            Height          =   270
            Left            =   105
            TabIndex        =   29
            Top             =   1590
            Width           =   2160
            _ExtentX        =   3810
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
            Caption         =   "Por Número de Cita"
         End
         Begin MSMask.MaskEdBox txtFCitaNew1 
            Height          =   315
            Left            =   990
            TabIndex        =   32
            Top             =   2460
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   12
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
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Sala"
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
            Index           =   9
            Left            =   645
            TabIndex        =   48
            Top             =   2160
            Width           =   315
         End
         Begin VB.Label lblReprogramacion 
            Alignment       =   1  'Right Justify
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   210
            Index           =   1
            Left            =   45
            TabIndex        =   46
            Top             =   1815
            Width           =   1755
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "N° Cita"
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
            Index           =   8
            Left            =   1920
            TabIndex        =   45
            Top             =   1815
            Width           =   570
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "F. nueva"
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
            Index           =   7
            Left            =   225
            TabIndex        =   44
            Top             =   2490
            Width           =   705
         End
         Begin VB.Label lblReprogramacion 
            AutoSize        =   -1  'True
            Caption         =   "PUnto Carga"
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
            Index           =   7
            Left            =   135
            TabIndex        =   37
            Top             =   285
            Width           =   1020
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "F. nueva Cita"
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
            Index           =   5
            Left            =   900
            TabIndex        =   36
            Top             =   1155
            Width           =   1065
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "F.Cita"
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
            Index           =   3
            Left            =   1515
            TabIndex        =   35
            Top             =   795
            Width           =   450
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Buscar Cita ya registrada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3225
         Left            =   30
         TabIndex        =   16
         Top             =   5040
         Width           =   3405
         Begin VB.CommandButton btnBuscar 
            Height          =   315
            Left            =   2025
            Picture         =   "ucSiListasCitas.ctx":37CC
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   615
            Width           =   1305
         End
         Begin VB.TextBox txtPacienteBuscar 
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
            Left            =   75
            MaxLength       =   99
            TabIndex        =   17
            Top             =   585
            Width           =   1920
         End
         Begin UltraGrid.SSUltraGrid grdCitas 
            Height          =   2175
            Left            =   45
            TabIndex        =   19
            Top             =   975
            Width           =   3315
            _ExtentX        =   5847
            _ExtentY        =   3836
            _Version        =   131072
            GridFlags       =   17040384
            LayoutFlags     =   67108884
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Lista de Citas"
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Apellido Paterno y Mat."
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
            Index           =   6
            Left            =   90
            TabIndex        =   18
            Top             =   330
            Width           =   1905
         End
      End
      Begin VB.Frame fraPtoCarga 
         Caption         =   "Actualiza Punto Carga"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3270
         Left            =   3240
         TabIndex        =   2
         Top             =   3300
         Width           =   3405
         Begin VB.Frame fraVarios 
            Height          =   825
            Left            =   30
            TabIndex        =   22
            Top             =   2385
            Width           =   3345
            Begin VB.CommandButton cmdReporgramacion 
               Caption         =   "Reprogramación"
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
               Left            =   45
               TabIndex        =   24
               Top             =   450
               Width           =   3285
            End
            Begin VB.CommandButton cmdDiasNoLaborables 
               Caption         =   "Días no Laborables x mes"
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
               Left            =   45
               TabIndex        =   23
               Top             =   105
               Width           =   3285
            End
         End
         Begin VB.TextBox txtCuposXdia 
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
            Height          =   330
            Left            =   2355
            MaxLength       =   99
            TabIndex        =   3
            Top             =   210
            Width           =   525
         End
         Begin VB.TextBox txtNroMinutos 
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
            Height          =   330
            Left            =   2355
            MaxLength       =   99
            TabIndex        =   4
            Top             =   555
            Width           =   525
         End
         Begin VB.CommandButton btnAceptar 
            Caption         =   "Actualizar"
            DisabledPicture =   "ucSiListasCitas.ctx":6415
            DownPicture     =   "ucSiListasCitas.ctx":6875
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   2070
            Picture         =   "ucSiListasCitas.ctx":6CEA
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   1260
            Width           =   1260
         End
         Begin MSMask.MaskEdBox txtHoraInicioCita 
            Height          =   315
            Left            =   2355
            TabIndex        =   15
            Top             =   900
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
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Hora de Inicio de Cita"
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
            Index           =   4
            Left            =   75
            TabIndex        =   14
            Top             =   900
            Width           =   1755
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "N° Máximo de  Citas por día"
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
            Index           =   1
            Left            =   75
            TabIndex        =   7
            Top             =   240
            Width           =   2250
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "N° minutos de atención"
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
            Index           =   2
            Left            =   75
            TabIndex        =   5
            Top             =   570
            Width           =   1950
         End
      End
      Begin MSDataListLib.DataList lstMedicos 
         Height          =   1320
         Left            =   60
         TabIndex        =   8
         ToolTipText     =   "Haga click sobre el nombre del médico para seleccionarlo"
         Top             =   405
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   2328
         _Version        =   393216
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
      Begin VB.Image ImgElegido 
         Height          =   195
         Left            =   3060
         Picture         =   "ucSiListasCitas.ctx":715F
         Top             =   180
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Puntos de Carga"
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
         TabIndex        =   9
         Top             =   210
         Width           =   1350
      End
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Citas"
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
      TabIndex        =   0
      Top             =   0
      Width           =   13320
   End
End
Attribute VB_Name = "ucSIcitasLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para lista de Programación Médica
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_Turnos() As doTurno
Dim mo_AdminProgramacionMedica As New SIGHNegocios.ReglasDeProgMedica
Dim mo_ReglasImagenes As New SIGHNegocios.ReglasImagenes
Dim mo_ReglasConfiguarcionReslab As New SIGHNegocios.ReglasConfiguarcionReslab
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_idUsuario As Long
Dim mb_SeHaModificadoProgramacion As Boolean
Dim ms_NombreUltimoMedicoSeleccionado As String
Dim mda_UltimaFechaSeleccionada As Date
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_cmbSalaNew As New sighentidades.ListaDespleglable

Public Event SePresionoTeclaEspecial(KeyCode As Integer)
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim lnMaximaHorasProgramadasXmedico As Integer
Dim oRsPuntosCarga As New Recordset
Dim ml_Area  As sghAreasLaboraEmpleado
Dim mi_IdCita As Long, ml_IdMovimiento As Long
Dim mi_idSala As Long, mi_IdPuntoCarga As Long
Dim ldHoy As Date
Dim lnNroCuposMaximosXDia As Long, lnIdResponsableNew1 As Long, lcTiempoAtencion11 As String

Private Const COLOR_CUPO_BLOQUEADO = &H80FFFF   'Amarillo
Private Const COLOR_CUPO_SEPARADO = &H1465FC   'naranja
Private Const COLOR_CUPO_DISPONIBLE = &HC0FFC0  'Verde
Private Const COLOR_CUPO_VENCIDO = &HD18D9C     'Morado
Private Const COLOR_DIA_PROGRAMADO = &HC0FFFF
Private Const COLOR_DIA_NO_PROGRAMADO = &HFFFFFF
Dim mda_UltimaMesSeleccionado As Integer
Dim mda_UltimaFechaSelecEnCalendarioFeriados As Date
Property Let Area(lValue As sghAreasLaboraEmpleado)
   ml_Area = lValue
End Property
Property Get Area() As sghAreasLaboraEmpleado
   Area = ml_Area
End Property


Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
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
'Property Let MenuAgregarEnabled(bValue As Boolean)
'   UserControl.mnuDiarioAgregarProgramacion.Enabled = bValue
'   UserControl.mnuCalAgregarProgramacion.Enabled = bValue
'End Property
'Property Let MenuModificarEnabled(bValue As Boolean)
'   UserControl.mnuDiarioModificarProgramacion.Enabled = bValue
'End Property
'Property Let MenuEliminarEnabled(bValue As Boolean)
'   UserControl.mnuDiarioEliminarProgramacion.Enabled = bValue
'   UserControl.mnuCalEliminarProgSelecionada.Enabled = bValue
'End Property
'Property Let MenuConsultarEnabled(bValue As Boolean)
'   UserControl.mnuDiarioConsultarProgramacion.Enabled = bValue
'End Property

Private Sub btnAceptar_Click()
    If Val(txtCuposXdia.Text) <= 0 Then
       MsgBox "El Nro CUPOS debe ser mayor a CERO", vbInformation, ""
       Exit Sub
    End If
    If Val(txtNroMinutos.Text) <= 0 Then
       MsgBox "El Nro MINUTOS x CUPO debe ser mayor a CERO", vbInformation, ""
       Exit Sub
    End If
    If sighentidades.EsHora(txtHoraInicioCita.Text) = False Then
       MsgBox "Verifique que la hora es correcta (00:00 a 23:00)", vbInformation, ""
       Exit Sub
    End If
    mo_reglasComunes.FactPuntosCargaActualizaCupos mi_IdPuntoCarga, Val(txtCuposXdia.Text), _
                      Val(txtNroMinutos.Text), txtHoraInicioCita.Text
     MsgBox "Se actualizó correctamente", vbInformation, ""
End Sub

Private Sub btnBuscar_Click()
    If txtPacienteBuscar.Text <> "" Then
       Dim oRsTmp1 As New Recordset
       Set oRsTmp1 = mo_ReglasImagenes.SiCitasFiltroPorPaciente(txtPacienteBuscar.Text)
       Set grdCitas.DataSource = oRsTmp1
    Else
       Set grdCitas.DataSource = Nothing
    End If
End Sub


Private Sub btnCancelar1_Click()
    fraPtoCarga.Visible = True
    fraReprogramacionCita.Visible = False
    fraDiasNoLaborables.Visible = False

End Sub



Private Sub Calendario_Change(ByVal NewDate As Date)
   
    'Si cambia de mes o año pregunta a guardar los datos
    'If Month(mda_UltimaFechaSeleccionada) <> Month(NewDate) Or Year(mda_UltimaFechaSeleccionada) <> Year(NewDate) Then
        'If mb_SeHaModificadoProgramacion Then
        '    If MsgBox("Ud ha modificado la programación del médico " + Chr(13) + UCase(ms_NombreUltimoMedicoSeleccionado) + Chr(13) + ", si no guarda los cambios se perderán. " + Chr(13) + "¿Desea guardar esos cambios? ", vbExclamation + vbYesNo, "Programación médica") = vbYes Then
        '        GrabarProgramacionDelMes
        '    End If
        '    mb_SeHaModificadoProgramacion = False
        'End If
        LimpiarProgramaciones
        LeerProgramacionDelMes Val(lstMedicos.BoundText), Month(NewDate), Year(NewDate)
    'End If
    
    mda_UltimaFechaSeleccionada = NewDate
    
    CargaProgramacionDeFechaElegida True, mda_UltimaFechaSeleccionada
    
    'mgaray COmentado para permitir elegir fechas dispersas
    'Diario.CurrentDate = NewDate
    Diario.Caption = Format(NewDate, "dddd, MMMM dd, yyyy")
    
    If lstMedicos.BoundText <> "" Then
       grdResumenCpt.Caption = "Resumen de CPT del día " & mda_UltimaFechaSeleccionada
       Set grdResumenCpt.DataSource = mo_ReglasImagenes.SiCitasResumenPorDia(Val(lstMedicos.BoundText), mda_UltimaFechaSeleccionada)
    Else
       Set grdResumenCpt.DataSource = Nothing
    End If
    txtFechaCita.Text = Format(mda_UltimaFechaSeleccionada, sighentidades.DevuelveFechaSoloFormato_DMY)
    If mda_UltimaMesSeleccionado <> Month(NewDate) Then
       mda_UltimaMesSeleccionado = Month(NewDate)
       LlenaCalendarioFeriados False, mda_UltimaFechaSeleccionada
    End If
End Sub

Private Sub Calendario_DateDblClick(ByVal DateClicked As Date)
'    Dim oProgInf As New ProgramacionInfDiaria
'
'    Set oProgInf.Diario = Diario
'    Set oProgInf.Calendario = Calendario
'    oProgInf.Show 1

End Sub

Private Sub Calendario_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        
    If Button = 2 Then
       ' PopupMenu mnuCalendario
    End If

End Sub



Sub RefrescarListaMedicos()
    Dim rsIdAlmacen As New Recordset
    Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
    Dim lnIdPuntoCargaDelUsuario As Long
    lnIdPuntoCargaDelUsuario = 0
    Set rsIdAlmacen = mo_AdminServiciosComunes.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(IIf(ml_Area = sghImageneología, sghImageneología, sghLaboratorio), sighentidades.Usuario)
    If rsIdAlmacen.RecordCount > 0 Then
       lnIdPuntoCargaDelUsuario = rsIdAlmacen!idLaboraSubArea
    End If
    rsIdAlmacen.Close
    Set rsIdAlmacen = Nothing
    Set mo_AdminServiciosComunes = Nothing

    If oRsPuntosCarga.State = 1 Then
       Set oRsPuntosCarga = Nothing
    End If
    With oRsPuntosCarga
        .Fields.Append "idGrupo", adInteger
        .Fields.Append "NombreGrupo", adVarChar, 50, adFldIsNullable
        .LockType = adLockOptimistic
        .Open
    End With
    Dim oRsTmpSalas As New Recordset
    Dim lcFiltro1 As String
    Set oRsTmpSalas = mo_ReglasImagenes.SiCitasSalasSeleccionarTodas
    If lnIdPuntoCargaDelUsuario > 0 Then
       lcFiltro1 = "idPuntoCarga=" & Trim(Str(lnIdPuntoCargaDelUsuario))
    Else
       lcFiltro1 = IIf(ml_Area = sghAreasLaboraEmpleado.sghImageneología, _
                   "(idPuntoCarga=20 or idPuntoCarga=23 or idPuntoCarga=21 or IdPuntoCarga=22)", _
                   "(idPuntoCarga=3 or idPuntoCarga=11 or idPuntoCarga=2)")
    End If
    oRsTmpSalas.Filter = lcFiltro1
    If oRsTmpSalas.RecordCount > 0 Then
        oRsTmpSalas.MoveFirst
        Do While Not oRsTmpSalas.EOF
            oRsPuntosCarga.AddNew
            oRsPuntosCarga!idGrupo = oRsTmpSalas!idSala
            oRsPuntosCarga!nombreGrupo = oRsTmpSalas!Sala
            oRsPuntosCarga.Update
            oRsTmpSalas.MoveNext
        Loop
    End If
    oRsTmpSalas.Close
    Set oRsTmpSalas = Nothing
      
    lstMedicos.BoundColumn = "IdGrupo"
    lstMedicos.ListField = "NombreGrupo"
    Set lstMedicos.RowSource = oRsPuntosCarga
    lstMedicos.Tag = ""

End Sub








Private Sub Calendario_NewMonth()
    LlenaCalendarioFeriados False, Calendario.Value
End Sub

Private Sub Calendario_NewYear()
    LlenaCalendarioFeriados False, Calendario.Value
End Sub



Private Sub cmbSalaNew_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbSalaNew
End Sub

Private Sub cmdActualizaDiasNoLaborables_Click()
    If MsgBox("Está seguro ?", vbQuestion + vbYesNo, "") <> vbNo Then
        Dim ldFechaInicial As Date, ldFechaFinal, lnMes As Integer, lnAnio As Integer
        Dim lcFiltro As String
        Dim oConexion As New Connection
        lnMes = Month(mda_UltimaFechaSelecEnCalendarioFeriados)
        lnAnio = Year(mda_UltimaFechaSelecEnCalendarioFeriados)
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        If mo_ReglasImagenes.SiFechasNoLaborablesEliminaXmesAnio(mi_idSala, lnMes, lnAnio, oConexion) Then
            ldFechaInicial = CDate("01/" & Right("0" & Trim(Str(lnMes)), 2) & "/" & Trim(Str(lnAnio)))
            ldFechaFinal = CDate(Format(sighentidades.DevuelveFechaHoraFinalDelMesDelMovimiento(mda_UltimaFechaSelecEnCalendarioFeriados), sighentidades.DevuelveFechaSoloFormato_DMY))
            Do While ldFechaInicial <= ldFechaFinal
               If Not PVCalendarFeriados.DATEImage(ldFechaInicial) Is Nothing Then
                  If mo_ReglasImagenes.SiFechasNoLaborablesAgregaFeriado(mi_idSala, ldFechaInicial, oConexion) Then
                  End If
               End If
               ldFechaInicial = ldFechaInicial + 1
            Loop
            oConexion.Close
            Set oConexion = Nothing
            LlenaCalendarioFeriados False, mda_UltimaFechaSeleccionada
            cmdSalir3_Click
            'Calendario_Change mda_UltimaFechaSeleccionada
        Else
           MsgBox "No se pudo GRABAR", vbInformation, ""
           oConexion.Close
           Set oConexion = Nothing
        End If
    End If
End Sub

Private Sub cmdDiasNoLaborables_Click()
    fraPtoCarga.Visible = False
    fraReprogramacionCita.Visible = False
    fraDiasNoLaborables.Visible = True
    lblReprogramacion(0).Caption = fraPtoCarga.Caption
    mda_UltimaFechaSelecEnCalendarioFeriados = mda_UltimaFechaSeleccionada
    PVCalendarFeriados.Value = mda_UltimaFechaSelecEnCalendarioFeriados
    LlenaCalendarioFeriados True, mda_UltimaFechaSelecEnCalendarioFeriados
    
End Sub

Sub ReprogramaPorNcita()
    If Not IsDate(txtFCitaNew1.Text) Then
       MsgBox "No es una FECHA " & txtFCitaNew1.Text, vbInformation, ""
       Exit Sub
    End If
    If lblReprogramacion(1).Caption = "" Then
       MsgBox "Esa CITA no tiene PACIENTE", vbInformation, ""
       Exit Sub
    End If
    If CDate(txtFCitaNew1.Text) < ldHoy Then
       MsgBox "La fecha " & txtFCitaNew1.Text & " tiene que ser HOY o mayor", vbInformation, ""
       Exit Sub
    End If
    If EsUnDiaFeriado(CDate(txtFCitaNew1.Text)) = True Then
       MsgBox "La fecha " & txtFCitaNew1.Text & " es DIA NO LABORABLE", vbInformation, ""
       Exit Sub
    End If
    If MsgBox("Está seguro ?", vbQuestion + vbYesNo, "") <> vbNo Then
       Dim oRsTmp1 As New Recordset
       Dim oRsTmp2 As New Recordset
       Dim oDoSiCitas As New DoSiCitas
       Dim oSiCitas As New SiCitas
       Dim oConexion As New Connection
       Dim mo_Procesos As New SIGHProxies.Procesos
       Dim lcHoraInicio1 As String, lcHoraFinal1 As String, lnCupos As Long
       Dim ldFechaNuevaCita As Date, lnIdResponsable As Long, lnIdProgramacion As Long
       lcHoraInicio1 = txtHoraPac.Text
       lcHoraFinal1 = mo_AdminProgramacionMedica.CalculaHoraFinal(txtHoraPac.Text, Val(lcTiempoAtencion11))
       Set oRsTmp1 = mo_ReglasImagenes.SiProgramacionXsalaFecha(Val(mo_cmbSalaNew.BoundText), CDate(txtFCitaNew1.Text))
       oRsTmp1.Filter = "HoraInicio<='" & lcHoraInicio1 & "' and HoraFin>='" & lcHoraInicio1 & "'"
       If oRsTmp1.RecordCount > 0 Then
          lnIdResponsable = oRsTmp1!idResponsable
          lnIdProgramacion = oRsTmp1!IdProgramacion
          oRsTmp1.Close
          oConexion.CursorLocation = adUseClient
          oConexion.CommandTimeout = 300
          oConexion.Open sighentidades.CadenaConexion
          Set oRsTmp1 = mo_ReglasImagenes.siCitasXidprogramacion(lnIdProgramacion, oConexion)
          lnCupos = oRsTmp1.RecordCount
          If lnCupos > 0 Then
             lnCupos = oRsTmp1!Cupo + 1
          Else
             lnCupos = 1
          End If
          oRsTmp1.Close
          Set oSiCitas.Conexion = oConexion
          oDoSiCitas.idCitaSI = Val(txtNCita1.Text)
          oDoSiCitas.IdUsuarioAuditoria = sighentidades.Usuario
          If oSiCitas.SeleccionarPorId(oDoSiCitas) Then
                oDoSiCitas.fecha = CDate(txtFCitaNew1.Text)
                oDoSiCitas.HoraInicio = lcHoraInicio1
                oDoSiCitas.HoraFinal = lcHoraFinal1
                oDoSiCitas.Cupo = lnCupos
                oDoSiCitas.idResponsable = lnIdResponsable
                oDoSiCitas.IdProgramacion = lnIdProgramacion
                oDoSiCitas.idSala = Val(mo_cmbSalaNew.BoundText)
                oDoSiCitas.llaveTicket = Format(Now, "ddmmyyhhmmss") & oDoSiCitas.HoraInicio & Trim(Str(oDoSiCitas.idSala))
                If oSiCitas.Modificar(oDoSiCitas) Then
                    If mo_ReglasImagenes.SiCitasDetalleActualizaLlave(oDoSiCitas.idCitaSI, oDoSiCitas.llaveTicket, oConexion) Then
                       mo_Procesos.EnviaMensajeCelularPorCuenta oDoSiCitas.idCuentaAtencion, "Se REPROGRAMO Cita N° " & _
                             oDoSiCitas.idCitaSI & " para el " & oDoSiCitas.fecha & " en " & lstMedicos.Text, "SiCita"
                    End If
                End If
          Else
                MsgBox "No existe PACIENTE para esa N° CITA", vbInformation, ""
          End If
          oConexion.Close
       Else
          oRsTmp1.Close
       End If
       Set oDoSiCitas = Nothing
       Set oSiCitas = Nothing
       Set oConexion = Nothing
       Set oRsTmp1 = Nothing
       Set oRsTmp2 = Nothing
       Set mo_Procesos = Nothing
       btnCancelar1_Click
       Calendario_Change mda_UltimaFechaSeleccionada
    End If
End Sub

Private Sub cmdGrabaReprogramacion_Click()
    If optPorFecha.Value = False Then
       ReprogramaPorNcita
       Exit Sub
    End If
    If Not IsDate(txtFechaCita.Text) Then
       MsgBox "No es una FECHA " & txtFechaCita.Text, vbInformation, ""
       Exit Sub
    End If
    If Not IsDate(txtFCitaNueva.Text) Then
       MsgBox "No es una FECHA " & txtFCitaNueva.Text, vbInformation, ""
       Exit Sub
    End If
    If CDate(txtFechaCita.Text) = CDate(txtFCitaNueva.Text) Then
       MsgBox "Deben ser FECHAS distintas", vbInformation, ""
       Exit Sub
    End If
    If CDate(txtFCitaNueva.Text) < ldHoy Then
       MsgBox "La fecha " & txtFCitaNueva.Text & " tiene que ser HOY o mayor", vbInformation, ""
       Exit Sub
    End If
    If EsUnDiaFeriado(CDate(txtFCitaNueva.Text)) = True Then
       MsgBox "La fecha " & txtFCitaNueva.Text & " es DIA NO LABORABLE", vbInformation, ""
       Exit Sub
    End If
    If MsgBox("Está seguro ?", vbQuestion + vbYesNo, "") <> vbNo Then
       Dim oRsTmp1 As New Recordset
       Dim oRsTmp2 As New Recordset
       Dim oDoSiCitas As New DoSiCitas
       Dim oSiCitas As New SiCitas
       Dim oConexion As New Connection
       Dim mo_Procesos As New SIGHProxies.Procesos
       Set oRsTmp1 = mo_ReglasImagenes.SiCitasPorDia(mi_idSala, CDate(txtFechaCita.Text))
       If oRsTmp1.RecordCount > 0 Then
          Set oRsTmp2 = mo_ReglasImagenes.SiCitasPorDia(mi_idSala, CDate(txtFCitaNueva.Text))
          If oRsTmp2.RecordCount = 0 Then
             oRsTmp2.Close
             Set oRsTmp2 = mo_ReglasImagenes.SiProgramacionXsalaFecha(mi_idSala, CDate(txtFCitaNueva.Text))
             If oRsTmp2.RecordCount = 0 Then
                MsgBox "No hay RESPONSABLE programados para " & txtFCitaNueva.Text & " " & oRsTmp1!HoraInicio, vbInformation, ""
             Else
                oConexion.CommandTimeout = 900
                oConexion.CursorLocation = adUseClient
                oConexion.Open sighentidades.CadenaConexion
                Set oSiCitas.Conexion = oConexion
                oRsTmp1.MoveFirst
                Do While Not oRsTmp1.EOF
                   oRsTmp2.Filter = "HoraInicio<='" & oRsTmp1!HoraInicio & "' and HoraFin>='" & oRsTmp1!HoraInicio & "'"
                   If oRsTmp2.RecordCount > 0 Then
                        oRsTmp2.MoveFirst
                        oDoSiCitas.idCitaSI = oRsTmp1!idCitaSI
                        oDoSiCitas.IdUsuarioAuditoria = sighentidades.Usuario
                        If oSiCitas.SeleccionarPorId(oDoSiCitas) Then
                           oDoSiCitas.idResponsable = oRsTmp2!idResponsable
                           oDoSiCitas.IdProgramacion = oRsTmp2!IdProgramacion
                           oDoSiCitas.fecha = CDate(txtFCitaNueva.Text)
                           If oSiCitas.Modificar(oDoSiCitas) Then
                              mo_Procesos.EnviaMensajeCelularPorCuenta oDoSiCitas.idCuentaAtencion, "Se REPROGRAMO Cita N° " & _
                                       oDoSiCitas.idCitaSI & " para el " & oDoSiCitas.fecha & " en " & lstMedicos.Text, "SiCita"
                           End If
                        End If
                   End If
                   oRsTmp1.MoveNext
                Loop
                oConexion.Close
                lstMedicos_Click
             End If
          Else
             MsgBox "Ya hay CITAS registradas en " & txtFCitaNueva.Text, vbInformation, ""
          End If
       Else
          MsgBox "No hay CITAS registradas en " & txtFechaCita.Text, vbInformation, ""
       End If
       Set oDoSiCitas = Nothing
       Set oSiCitas = Nothing
       Set oConexion = Nothing
       Set oRsTmp1 = Nothing
       Set oRsTmp2 = Nothing
       Set mo_Procesos = Nothing
       btnCancelar1_Click
       Calendario_Change mda_UltimaFechaSeleccionada
    End If
End Sub

Private Sub cmdReporgramacion_Click()
    fraPtoCarga.Visible = False
    fraReprogramacionCita.Visible = True
    fraDiasNoLaborables.Visible = False
    lblReprogramacion(7).Caption = fraPtoCarga.Caption
    txtFCitaNueva.Text = sighentidades.FECHA_VACIA_DMY
    txtNCita1.Text = ""
    lblReprogramacion(1).Caption = ""
    txtFCitaNew1.Text = sighentidades.FECHA_VACIA_DMY
End Sub







Private Sub cmdSalir3_Click()
    fraPtoCarga.Visible = True
    fraReprogramacionCita.Visible = False
    fraDiasNoLaborables.Visible = False

End Sub



Private Sub Diario_Click()
    On Error Resume Next
    Dim programacion As PVAppointment
    
    Set programacion = Diario.AppointmentSet.GetSelectedAppointment
    mi_IdCita = programacion.DataVariant.idCitaSI
    ml_IdMovimiento = programacion.DataVariant.IdMovimiento
End Sub

Private Sub Diario_KeyPress(ByVal KeyAscii As Integer)
'    If KeyAscii = 13 Then
'       ucLaborPacienConCitas1.Visible = True
'       ucLaborPacienConCitas1.LlenaCitasPorFecha mda_UltimaFechaSeleccionada
'    End If
End Sub

Private Sub Diario_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then
        'PopupMenu mnuDiario
    End If
End Sub

Private Sub grdCitas_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
     grdCitas.Bands(0).Columns("paciente").Width = 2000
End Sub

Private Sub grdResumenCpt_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
     grdResumenCpt.Bands(0).Columns("codigo").Width = 600
     grdResumenCpt.Bands(0).Columns("nombre").Width = 2400
     grdResumenCpt.Bands(0).Columns("cant").Width = 400
     mo_Apariencia.ConfigurarFilasBiColores grdResumenCpt, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub lstMedicos_Click()
    
    'Verifica que no sea el mismo medico
    If lstMedicos.Tag = lstMedicos.BoundText Then
        Exit Sub
    End If
    mi_idSala = Val(lstMedicos.BoundText)
    
    PuntoCargaDatos
    LimpiarProgramaciones
    LeerProgramacionDelMes Val(lstMedicos.BoundText), Month(Diario.CurrentDate), Year(Diario.CurrentDate)
    
    lstMedicos.Tag = lstMedicos.BoundText
    ms_NombreUltimoMedicoSeleccionado = lstMedicos.Text
    PuntoCargaDatos
    fraPtoCarga.Caption = lstMedicos.Text
    
    'mi_idSala = Val(lstMedicos.BoundText)
    cmdSalir3_Click
    LlenaCalendarioFeriados False, mda_UltimaFechaSeleccionada
    
End Sub

Sub CambiaHoraInicioDelDiario()
'    If txtHoraInicioCita.Text <> SIGHEntidades.HORA_VACIA_HM Then
'       If txtHoraInicioCita.Text > "12:59" Then
'          Diario.BusinessHoursBegin = txtHoraInicioCita.Text & " PM"
'       Else
'          Diario.BusinessHoursBegin = txtHoraInicioCita.Text & " AM"
'       End If
'    End If
End Sub

Function EsUnDiaFeriado(ldFechaSeleccionada As Date) As Boolean
    EsUnDiaFeriado = False
    Dim oRsTmp1 As New Recordset
    Set oRsTmp1 = mo_ReglasImagenes.SiFechasNoLaborablesXmesAnio(mi_idSala, Month(ldFechaSeleccionada), Year(ldFechaSeleccionada))
    oRsTmp1.Filter = "fechaNoLaborable='" & Format(ldFechaSeleccionada, sighentidades.DevuelveFechaSoloFormato_DMY) & "'"
    If oRsTmp1.RecordCount > 0 Then
       EsUnDiaFeriado = True
    End If
    oRsTmp1.Close
    Set oRsTmp1 = Nothing
End Function

Sub LlenaCalendarioFeriados(lbSoloCalendarioDeFeriados As Boolean, ldFechaSeleccionada As Date)
    Dim ldFechaInicial As Date, ldFechaFinal
    Dim lcFiltro As String
    Dim oRsTmp1 As New Recordset
    If ldFechaSeleccionada = 0 Then
       ldFechaSeleccionada = Date
    End If
    Set oRsTmp1 = mo_ReglasImagenes.SiFechasNoLaborablesXmesAnio(mi_idSala, Month(ldFechaSeleccionada), Year(ldFechaSeleccionada))
    If oRsTmp1.RecordCount > 0 Then
        ldFechaInicial = CDate("01/" & Right("0" & Trim(Str(Month(ldFechaSeleccionada))), 2) & "/" & Trim(Str(Year(ldFechaSeleccionada))))
        ldFechaFinal = CDate(Format(sighentidades.DevuelveFechaHoraFinalDelMesDelMovimiento(ldFechaSeleccionada), sighentidades.DevuelveFechaSoloFormato_DMY))
        
        Do While ldFechaInicial <= ldFechaFinal
'If Day(ldFechaInicial) = 19 Then
'lcFiltro = ""
'End If
           If lbSoloCalendarioDeFeriados = False Then
              Calendario.DATEImage(ldFechaInicial) = Nothing
           End If
           PVCalendarFeriados.DATEImage(ldFechaInicial) = Nothing
           oRsTmp1.Filter = "fechaNoLaborable='" & Format(ldFechaInicial, sighentidades.DevuelveFechaSoloFormato_DMY) & "'"
           If oRsTmp1.RecordCount > 0 Then
                If lbSoloCalendarioDeFeriados = False Then
                   Calendario.DATEImage(ldFechaInicial) = ImgElegido.Picture
                End If
                PVCalendarFeriados.DATEImage(ldFechaInicial) = ImgElegido.Picture
           End If
           ldFechaInicial = ldFechaInicial + 1
        Loop
    End If
    oRsTmp1.Close
    Set oRsTmp1 = Nothing
End Sub

Sub PuntoCargaDatos()
    Dim oRsTmpSalas As New Recordset
    Set oRsTmpSalas = mo_ReglasImagenes.SiCitasSalasSeleccionarTodas
    oRsTmpSalas.Filter = "idSala=" & lstMedicos.BoundText
    mi_IdPuntoCarga = oRsTmpSalas!idPuntoCarga
    oRsTmpSalas.Close
    Set oRsTmpSalas = Nothing
    
    txtCuposXdia.Text = ""
    txtNroMinutos.Text = ""
    txtHoraInicioCita.Text = sighentidades.HORA_VACIA_HM
    Dim oRsTmp9 As New Recordset
    Set oRsTmp9 = mo_reglasComunes.FactPuntosCargaSeleccionarPorId(mi_IdPuntoCarga)
    If oRsTmp9.RecordCount > 0 Then
       txtCuposXdia.Text = IIf(IsNull(oRsTmp9!NroCupos), 0, oRsTmp9!NroCupos)
       txtNroMinutos.Text = IIf(IsNull(oRsTmp9!nroCuposMinutos), 0, oRsTmp9!nroCuposMinutos)
       txtHoraInicioCita.Text = IIf(IsNull(oRsTmp9!HoraInicioDiaCita), sighentidades.HORA_VACIA_HM, oRsTmp9!HoraInicioDiaCita)
       CambiaHoraInicioDelDiario
    End If
    oRsTmp9.Close
    Set oRsTmp9 = Nothing
End Sub

Private Sub mnuCalAgregarProgramacion_Click()
'Dim oProgDetalle As New LaboratorioProgDetalle
'
'    If lstMedicos.BoundText = "" Then
'        MsgBox "Seleccione un GRUPO EXAMEN", vbInformation, "Programación médica"
'        Exit Sub
'    End If
'
'
'    oProgDetalle.FechaInicial = mda_UltimaFechaSeleccionada
'    oProgDetalle.IdProgramacion = 0
'    oProgDetalle.idGrupo = Val(lstMedicos.BoundText)
'    oProgDetalle.idUsuario = Me.idUsuario
'    oProgDetalle.Opcion = sghAgregar
'    oProgDetalle.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
'    oProgDetalle.lcNombrePc = mo_lcNombrePc
'    oProgDetalle.Show 1
'
'
'    LimpiarProgramaciones
'    'Dim lnIdGrupo9 As Long
'    'lnIdGrupo9 = oProgDetalle.idGrupo
'    LeerProgramacionDelMes oProgDetalle.idGrupo, Month(Diario.CurrentDate), Year(Diario.CurrentDate)
'    Unload oProgDetalle
    
End Sub




Public Sub mnuDiarioConsultarProgramacion_Click()
    If mda_UltimaFechaSeleccionada = 0 Then
       MsgBox "Elija la FECHA", vbInformation, ""
       Exit Sub
    End If
    If Val(lstMedicos.BoundText) = 0 Then
       MsgBox "Elija el PUNTO DE CARGA", vbInformation, ""
       Exit Sub
    End If
    If ml_Area = sghAreasLaboraEmpleado.sghImageneología Then
        Dim oSiCitaDetalleIMG As New SiCitaDetalleIMG
        oSiCitaDetalleIMG.idSala = mi_idSala
        oSiCitaDetalleIMG.IdMovimiento = mi_IdCita
        oSiCitaDetalleIMG.fechaCita = mda_UltimaFechaSeleccionada
        oSiCitaDetalleIMG.PuntoCarga = Val(lstMedicos.BoundText)
        oSiCitaDetalleIMG.Opcion = sghConsultar
        oSiCitaDetalleIMG.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
        oSiCitaDetalleIMG.lcNombrePc = mo_lcNombrePc
        oSiCitaDetalleIMG.Show 1
        Unload oSiCitaDetalleIMG
    Else
        Dim oSiCitaDetalle As New SiCitaDetalle
        oSiCitaDetalle.IdMovimiento = mi_IdCita
        oSiCitaDetalle.fechaCita = mda_UltimaFechaSeleccionada
        oSiCitaDetalle.PuntoCarga = Val(lstMedicos.BoundText)
        oSiCitaDetalle.Opcion = sghConsultar
        oSiCitaDetalle.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
        oSiCitaDetalle.lcNombrePc = mo_lcNombrePc
        oSiCitaDetalle.Show 1
        Unload oSiCitaDetalle
    End If
    LimpiarProgramaciones
    LeerProgramacionDelMes Val(lstMedicos.BoundText), Month(Diario.CurrentDate), Year(Diario.CurrentDate)
    
End Sub



Private Sub optPorFecha_Click(Value As Integer)
    lnIdResponsableNew1 = 0
End Sub

Private Sub optPorNcita_Click(Value As Integer)
    lnIdResponsableNew1 = 0
End Sub

Private Sub PVCalendarFeriados_Change(ByVal NewDate As Date)
   
    mda_UltimaFechaSelecEnCalendarioFeriados = NewDate
    If PVCalendarFeriados.DATEImage(NewDate) Is Nothing Then
        PVCalendarFeriados.DATEImage(NewDate) = ImgElegido.Picture
    Else
        Set PVCalendarFeriados.DATEImage(NewDate) = Nothing
    End If
    'lblReprogramacion(0).Caption = mda_UltimaFechaSelecEnCalendarioFeriados
End Sub





Private Sub PVCalendarFeriados_NewMonth()
    LlenaCalendarioFeriados True, PVCalendarFeriados.Value
End Sub

Private Sub PVCalendarFeriados_NewYear()
    LlenaCalendarioFeriados True, PVCalendarFeriados.Value
End Sub

Private Sub txtCuposXdia_KeyDown(KeyCode As Integer, Shift As Integer)
        mo_Teclado.RealizarNavegacion KeyCode, txtCuposXdia
End Sub



Private Sub txtCuposXdia_KeyPress(KeyAscii As Integer)
            If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
                KeyAscii = 0
            End If
End Sub







Private Sub txtFCitaNew1_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtFCitaNew1
End Sub

Private Sub txtFCitaNew1_LostFocus()
    If cmbSalaNew.Text = "" Then
       MsgBox "Elija la SALA", vbInformation, ""
    ElseIf Not IsDate(txtFCitaNew1.Text) Then
       MsgBox "No es una FECHA correcta", vbInformation, ""
    Else
       LlenaComboConHoraInicio
    End If
End Sub
Sub LlenaComboConHoraInicio()
    If EsFecha(txtFCitaNew1.Text, "DD/MM/AAAA") = True And lnIdResponsableNew1 > 0 Then
       Dim oRsTmp1 As New Recordset
       Dim oRsTmp2 As New Recordset
       Dim lcHoraInicioCita As String, lbCitaVacia As Boolean, lnYaCitados As Long
       txtHoraPac.Clear
       Set oRsTmp1 = mo_ReglasImagenes.SiProgramacionXsalaFecha(Val(mo_cmbSalaNew.BoundText), CDate(txtFCitaNew1.Text))
       Set oRsTmp2 = mo_ReglasImagenes.SiCitasSeleccionarPorSalaYFecha(Val(mo_cmbSalaNew.BoundText), CDate(txtFCitaNew1.Text))
       lnYaCitados = oRsTmp2.RecordCount
       If oRsTmp1.RecordCount > 0 Then
          oRsTmp1.MoveFirst
          lcTiempoAtencion11 = Trim(Str(oRsTmp1!TiempoPromedioAtencion))
          Do While Not oRsTmp1.EOF
             lcHoraInicioCita = oRsTmp1!HoraInicio
             Do While True
                lbCitaVacia = True
                If lnYaCitados > 0 Then
                   oRsTmp2.MoveFirst
                   oRsTmp2.Find "horaInicio='" & lcHoraInicioCita & "'"
                   If Not oRsTmp2.EOF Then
                      lbCitaVacia = False
                   End If
                End If
                If lbCitaVacia = True Then
                   txtHoraPac.AddItem lcHoraInicioCita
                End If
                lcHoraInicioCita = mo_AdminProgramacionMedica.ConvertirAHora(mo_AdminProgramacionMedica.ConvertirAMinutos(lcHoraInicioCita) + Val(lcTiempoAtencion11))
                If lcHoraInicioCita >= oRsTmp1!HoraFin Then
                   Exit Do
                End If
             Loop
             oRsTmp1.MoveNext
          Loop
       End If
       oRsTmp1.Close
       Set oRsTmp1 = Nothing
       Set oRsTmp2 = Nothing
    End If
End Sub


Private Sub txtFCitaNueva_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtFCitaNueva
End Sub

Private Sub txtFechaCita_KeyDown(KeyCode As Integer, Shift As Integer)
        mo_Teclado.RealizarNavegacion KeyCode, txtFechaCita
    
End Sub





Private Sub txtNCita1_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtNCita1
End Sub

Private Sub txtNCita1_LostFocus()
   lnIdResponsableNew1 = 0
   If Val(txtNCita1.Text) > 0 Then
       txtHoraPac.Clear
       Dim oDoSiCitas As New DoSiCitas
       Dim oSiCitas As New SiCitas
       Dim oConexion As New Connection
       Dim oRsTmpSalas1 As New Recordset
       oConexion.CursorLocation = adUseClient
       oConexion.CommandTimeout = 300
       oConexion.Open sighentidades.CadenaConexion
       oDoSiCitas.idCitaSI = Val(txtNCita1.Text)
       oDoSiCitas.IdUsuarioAuditoria = sighentidades.Usuario
       Set oSiCitas.Conexion = oConexion
       If oSiCitas.SeleccionarPorId(oDoSiCitas) Then
          lblReprogramacion(1).Caption = oDoSiCitas.Paciente
          lnIdResponsableNew1 = oDoSiCitas.idResponsable
          
          Set oRsTmpSalas1 = mo_ReglasImagenes.SiCitasSalasSeleccionarTodas
          oRsTmpSalas1.Filter = "idPuntoCarga=" & Trim(Str(oDoSiCitas.idPuntoCarga))
          mo_cmbSalaNew.ListField = "Sala"
          mo_cmbSalaNew.BoundColumn = "IdSala"
          Set mo_cmbSalaNew.RowSource = oRsTmpSalas1
          If oRsTmpSalas1.RecordCount = 1 Then
             oRsTmpSalas1.MoveFirst
             mo_cmbSalaNew.BoundText = oRsTmpSalas1!idSala
          End If
          
          
       Else
          lblReprogramacion(1).Caption = ""
          MsgBox "No existe PACIENTE para esa N° CITA", vbInformation, ""
       End If
       Set oDoSiCitas = Nothing
       Set oSiCitas = Nothing
       Set oConexion = Nothing
       Set oRsTmpSalas1 = Nothing
   End If
End Sub

Private Sub txtNroMinutos_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtNroMinutos
End Sub

Private Sub txtNroMinutos_KeyPress(KeyAscii As Integer)
            If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
                KeyAscii = 0
            End If
End Sub

'Private Sub ucLaborPacienConCitas1_SePulsoClicEnSalir(KeyCode As Boolean)
'    If KeyCode = True Then
'       ucLaborPacienConCitas1.Visible = False
'    End If
'End Sub

Private Sub UserControl_Initialize()
    
    'Calendario.AttachDayView Diario
    'Diario.AttachCalendar Calendario
    
End Sub

Public Function Inicializar()
    Set mo_cmbSalaNew.MiComboBox = cmbSalaNew
    
    Calendario.AttachDayView Diario
    Diario.AttachCalendar Calendario
    
    ConfigurarMenusProgramacionMedica
    lnMaximaHorasProgramadasXmedico = Val(lcBuscaParametro.SeleccionaFilaParametro(309))
    
    RefrescarListaMedicos
    ConfiguraPermisos
    ldHoy = CDate(lcBuscaParametro.RetornaFechaServidorSQL)
    mo_Apariencia.ConfigurarFilasBiColores grdCitas, sighentidades.GrillaConFilasBicolor
End Function



Sub ConfiguraPermisos()
    'PERMISOS
    Dim oRsPermisos As New Recordset
    Set oRsPermisos = mo_ReglasSeguridad.UsuariosRolesSeleccionarPermisosTodos(sighentidades.Usuario)
    oRsPermisos.Filter = "idPermiso=550"  'Tiene permiso para modificar CUPOS de Pto.Carga
    btnAceptar.Visible = False
    If oRsPermisos.RecordCount > 0 Then
        btnAceptar.Visible = True
    Else
        mo_Formulario.HabilitarDeshabilitar txtCuposXdia, False
        mo_Formulario.HabilitarDeshabilitar txtNroMinutos, False
        mo_Formulario.HabilitarDeshabilitar txtHoraInicioCita, False
    End If
    oRsPermisos.Close
    Set oRsPermisos = Nothing
End Sub

Private Sub UserControl_Resize()
   lblNombre.Width = UserControl.Width
   fraProgramacion.Height = fraMedico.Height
   Diario.Height = fraProgramacion.Height - 900
   lstProgramacion.Top = Diario.Top + Diario.Height + 50
   Calendario.Width = fraProgramacion.Width - Diario.Width - 100 ' 4280
   grdResumenCpt.Width = fraProgramacion.Width - Diario.Width - 100 ' 4280
   
    fraPtoCarga.Top = 1770
    fraPtoCarga.Left = 45
    fraPtoCarga.Visible = True
    fraReprogramacionCita.Visible = False
    fraDiasNoLaborables.Visible = False
    fraReprogramacionCita.Top = fraPtoCarga.Top
    fraReprogramacionCita.Left = fraPtoCarga.Left
    fraReprogramacionCita.Height = fraPtoCarga.Height
    fraReprogramacionCita.Width = fraPtoCarga.Width
    fraDiasNoLaborables.Top = fraPtoCarga.Top
    fraDiasNoLaborables.Left = fraPtoCarga.Left
    fraDiasNoLaborables.Height = fraPtoCarga.Height
    fraDiasNoLaborables.Width = fraPtoCarga.Width
End Sub
Private Sub UserControl_Terminate()
'    Calendario.AttachDayView Nothing
'    Diario.AttachCalendar Nothing
End Sub
Public Sub ConfigurarMenusProgramacionMedica()


    

End Sub



Public Sub mnuDiarioAgregarProgramacion_Click()

    If mda_UltimaFechaSeleccionada = 0 Then
       MsgBox "Elija la FECHA", vbInformation, ""
       Exit Sub
    End If
    If Val(lstMedicos.BoundText) = 0 Then
       MsgBox "Elija la SALA", vbInformation, ""
       Exit Sub
    End If
    If ldHoy > Diario.CurrentDate Then
       MsgBox "No se puede Agrear CITA menor a HOY", vbInformation, "Asignacion de citas"
       Exit Sub
    End If
    If Not Calendario.DATEImage(mda_UltimaFechaSeleccionada) Is Nothing Then
       MsgBox "No se puede Agrear CITA porque es DIA NO LABORABLE", vbInformation, "Asignacion de citas"
       Exit Sub
    End If
    If ml_Area = sghAreasLaboraEmpleado.sghImageneología Then
        Dim oSiCitaDetalleIMG As New SiCitaDetalleIMG
        oSiCitaDetalleIMG.idSala = Val(lstMedicos.BoundText)
        oSiCitaDetalleIMG.fechaCita = mda_UltimaFechaSeleccionada
        oSiCitaDetalleIMG.PuntoCarga = mi_IdPuntoCarga
        oSiCitaDetalleIMG.Opcion = sghAgregar
        oSiCitaDetalleIMG.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
        oSiCitaDetalleIMG.lcNombrePc = mo_lcNombrePc
        oSiCitaDetalleIMG.Show 1
        Unload oSiCitaDetalleIMG
    Else
        Dim oSiCitaDetalle As New SiCitaDetalle
        oSiCitaDetalle.idSala = Val(lstMedicos.BoundText)
        oSiCitaDetalle.fechaCita = mda_UltimaFechaSeleccionada
        oSiCitaDetalle.PuntoCarga = mi_IdPuntoCarga
        oSiCitaDetalle.Opcion = sghAgregar
        oSiCitaDetalle.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
        oSiCitaDetalle.lcNombrePc = mo_lcNombrePc
        oSiCitaDetalle.Show 1
        Unload oSiCitaDetalle
    End If
    LimpiarProgramaciones
    LeerProgramacionDelMes Val(lstMedicos.BoundText), Month(Diario.CurrentDate), Year(Diario.CurrentDate)
End Sub

Public Sub mnuDiarioEliminarProgramacion_Click()
    If mda_UltimaFechaSeleccionada = 0 Then
       MsgBox "Elija la FECHA", vbInformation, ""
       Exit Sub
    End If
    If Val(lstMedicos.BoundText) = 0 Then
       MsgBox "Elija la SALA", vbInformation, ""
       Exit Sub
    End If
    If ldHoy > Diario.CurrentDate Then
       MsgBox "No se puede Eliminar CITA menor a HOY", vbInformation, "Asignacion de citas"
       Exit Sub
    End If
    If ml_IdMovimiento > 0 Then
       MsgBox "No se puede Eliminar CITA porque ya tiene MOVIMIENTO ", vbInformation, "Asignacion de citas"
       Exit Sub
    End If
    If ml_Area = sghAreasLaboraEmpleado.sghImageneología Then
        Dim oSiCitaDetalleIMG As New SiCitaDetalleIMG
        oSiCitaDetalleIMG.idSala = mi_idSala
        oSiCitaDetalleIMG.IdMovimiento = mi_IdCita
        oSiCitaDetalleIMG.fechaCita = mda_UltimaFechaSeleccionada
        oSiCitaDetalleIMG.PuntoCarga = mi_IdPuntoCarga
        oSiCitaDetalleIMG.Opcion = sghEliminar
        oSiCitaDetalleIMG.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
        oSiCitaDetalleIMG.lcNombrePc = mo_lcNombrePc
        oSiCitaDetalleIMG.Show 1
        Unload oSiCitaDetalleIMG
    Else
        Dim oSiCitaDetalle As New SiCitaDetalle
        oSiCitaDetalle.IdMovimiento = mi_IdCita
        oSiCitaDetalle.fechaCita = mda_UltimaFechaSeleccionada
        oSiCitaDetalle.PuntoCarga = mi_IdPuntoCarga
        oSiCitaDetalle.Opcion = sghEliminar
        oSiCitaDetalle.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
        oSiCitaDetalle.lcNombrePc = mo_lcNombrePc
        oSiCitaDetalle.Show 1
        Unload oSiCitaDetalle
    End If

    LimpiarProgramaciones
    LeerProgramacionDelMes Val(lstMedicos.BoundText), Month(Diario.CurrentDate), Year(Diario.CurrentDate)
    

End Sub

Public Sub mnuDiarioModificarProgramacion_Click()
    If mda_UltimaFechaSeleccionada = 0 Then
       MsgBox "Elija la FECHA", vbInformation, ""
       Exit Sub
    End If
    If Val(lstMedicos.BoundText) = 0 Then
       MsgBox "Elija la SALA", vbInformation, ""
       Exit Sub
    End If
    If ldHoy > Diario.CurrentDate Then
       MsgBox "No se puede Modificar CITA menor a HOY", vbInformation, "Asignacion de citas"
       Exit Sub
    End If
    If ml_IdMovimiento > 0 Then
       MsgBox "No se puede Modificar CITA porque ya tiene MOVIMIENTO: " & Trim(Str(ml_IdMovimiento)), vbInformation, "Asignacion de citas"
       Exit Sub
    End If
    If ml_Area = sghAreasLaboraEmpleado.sghImageneología Then
        Dim oSiCitaDetalleIMG As New SiCitaDetalleIMG
        oSiCitaDetalleIMG.idSala = mi_idSala
        oSiCitaDetalleIMG.IdMovimiento = mi_IdCita
        oSiCitaDetalleIMG.fechaCita = mda_UltimaFechaSeleccionada
        oSiCitaDetalleIMG.PuntoCarga = mi_IdPuntoCarga
        oSiCitaDetalleIMG.Opcion = sghModificar
        oSiCitaDetalleIMG.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
        oSiCitaDetalleIMG.lcNombrePc = mo_lcNombrePc
        oSiCitaDetalleIMG.Show 1
        Unload oSiCitaDetalleIMG
    Else
        Dim oSiCitaDetalle As New SiCitaDetalle
        oSiCitaDetalle.IdMovimiento = mi_IdCita
        oSiCitaDetalle.fechaCita = mda_UltimaFechaSeleccionada
        oSiCitaDetalle.PuntoCarga = mi_IdPuntoCarga
        oSiCitaDetalle.Opcion = sghModificar
        oSiCitaDetalle.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
        oSiCitaDetalle.lcNombrePc = mo_lcNombrePc
        oSiCitaDetalle.Show 1
        Unload oSiCitaDetalle
    End If

    LimpiarProgramaciones
    LeerProgramacionDelMes Val(lstMedicos.BoundText), Month(Diario.CurrentDate), Year(Diario.CurrentDate)
    
End Sub

Private Sub mnuCalEliminarProgSelecionada_Click()
'Dim daDiaSeleccionado As Date
'Dim programacion As PVAppointment
'Dim sTitulo As String
'Dim sHoras() As String
'Dim iHoraIni As Integer
'Dim iHoraFin As Integer
'Dim bTurnoProgramado As Boolean
'
'
'
'    daDiaSeleccionado = Calendario.Value
'    Do While daDiaSeleccionado <> 0
'        sTitulo = ""
'        Set programacion = Diario.AppointmentSet.Get(daDiaSeleccionado)
'
'        bTurnoProgramado = False
'        Do While Not programacion Is Nothing
'            'Verifica que la programacion sea del mismo dia
'            If Format(programacion.StartDateTime, sighentidades.DevuelveFechaSoloFormato_DMY) = daDiaSeleccionado Then
'                If Not mo_AdminProgramacionMedica.ProgramacionMedicaEliminar(programacion.DataVariant, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "") Then
'                    MsgBox mo_AdminProgramacionMedica.MensajeError, vbInformation, "Programación médica"
'                Else
'                    Diario.AppointmentSet.Remove programacion.Key
'                    Calendario.DATEText(daDiaSeleccionado) = ""
'                End If
'            Else
'                Exit Sub
'            End If
'            Set programacion = Diario.AppointmentSet.GetNext(programacion)
'        Loop
'        daDiaSeleccionado = Calendario.NextSelectedDate(daDiaSeleccionado)
'    Loop
'
'    mb_SeHaModificadoProgramacion = True
'
'    LimpiarProgramaciones
'    LeerProgramacionDelMes Val(lstMedicos.BoundText), Month(Diario.CurrentDate), Year(Diario.CurrentDate)
'
End Sub

Sub LeerProgramacionDelMes(lnIdPuntoCarga As Long, iMes As Integer, iAnio As Integer)
Dim oProgramaciones As Collection
Dim oProgramacion As DoSiCitas
Dim programacion As PVAppointment
Dim sHoras() As String
Dim iHoraIni As Double
Dim iHoraFin As Double
Dim daFechaIni As Date
Dim sDescripcion As String

Dim dCantidadHoras As Double
Dim dCantidadHorasMes As Double
Dim oRsTmp As New Recordset
Dim lcNombreServicio As String, bPrimeraCita As Boolean
Dim lcSql As String, lcFecha As String, lnCitasAsignadas As Integer
Dim iNroDiasMes As Integer, i As Integer, lcCitasYaAsignadas As String, ldFechaConCitas As Date
Dim lcHoraInicioProgramacion As String, lnIdSala99 As Long
        '
        lcCitasYaAsignadas = ""
        lnCitasAsignadas = 0
        'Obtiene las programaciones del medico del mes correspondiente
        Set oProgramaciones = mo_ReglasImagenes.SiCitasLeerPorMedicoYMes(lnIdPuntoCarga, iMes, iAnio)
        If oProgramaciones.Count > 0 Then
            bPrimeraCita = True
            daFechaIni = oProgramaciones.Item(1).fecha
            ldFechaConCitas = oProgramaciones.Item(1).fecha
            lcHoraInicioProgramacion = oProgramaciones.Item(1).HoraInicio
            lnIdSala99 = oProgramaciones.Item(1).idSala
            Set oRsTmp = mo_ReglasImagenes.siProgramacionXsalaYfecha(lnIdSala99, ldFechaConCitas)
            If oRsTmp.RecordCount > 0 Then
               lcHoraInicioProgramacion = oRsTmp!HoraInicio
            End If
            oRsTmp.Close
            
            For Each oProgramacion In oProgramaciones
                If bPrimeraCita = True Then
                   bPrimeraCita = False
                   Diario.BusinessHoursBegin = Format(lcHoraInicioProgramacion, sighentidades.DevuelveHoraSoloFormato_HMS)
                End If

                oProgramacion.IdUsuarioAuditoria = ml_idUsuario
                'Agrega programacion
                sHoras = Split(oProgramacion.HoraInicio, ":")
                iHoraIni = Val(sHoras(0)) + IIf(Val(sHoras(1)) = 59, 60, Val(sHoras(1))) / 60

                sHoras = Split(oProgramacion.HoraFinal, ":")
                iHoraFin = Val(sHoras(0)) + IIf(Val(sHoras(1)) = 59, 60, Val(sHoras(1))) / 60

                dCantidadHoras = Format(iHoraFin - iHoraIni, "##0.00")
                dCantidadHorasMes = dCantidadHorasMes + dCantidadHoras   'Val(txtNroMinutos.Text)
                'busca Servicio
                lcFecha = Calendario.Value
                lcNombreServicio = oProgramacion.Paciente



                Set programacion = Diario.AppointmentSet.Add("N° Cupo: " & Trim(Str(oProgramacion.Cupo)) & _
                                   " - " & lcNombreServicio & " - N° Cita: " & Trim(Str(oProgramacion.idCitaSI)) & _
                                   IIf(oProgramacion.IdMovimiento > 0, " - Mov: " & Trim(Str(oProgramacion.IdMovimiento)), ""), _
                                   oProgramacion.fecha + iHoraIni / 24, oProgramacion.fecha + iHoraFin / 24)
                programacion.DataVariant = oProgramacion
                programacion.ReadOnly = True
                sDescripcion = "(" & txtNroMinutos.Text & ")"
                If daFechaIni = oProgramacion.fecha Then
                    'Si hay mas de una programación en la misma fecha, concatena los códigos
                    'sDescripcion = sDescripcion + IIf(sDescripcion <> "", "/", "") & "(" & dCantidadHoras & ")"
                Else
                    'Si es la primera programación en el dia
                    Calendario.DATEText(daFechaIni) = Trim(Str(lnCitasAsignadas))
                    Calendario.DATEBackColor(daFechaIni) = vbBlue
                    'sDescripcion = "(" & oProgramacion.cuposCE & ")"
                    daFechaIni = oProgramacion.fecha
                    Calendario.DATEForeColor(daFechaIni) = vbBlack
                    If oProgramacion.IdEstado <> 1 Then
                       Calendario.DATEForeColor(daFechaIni) = vbRed
                    End If
                End If
                If ldFechaConCitas = oProgramacion.fecha Then
                   lnCitasAsignadas = lnCitasAsignadas + 1
                Else
                   If lnCitasAsignadas > 0 Then
                        lcCitasYaAsignadas = lcCitasYaAsignadas & "/" & Trim(Str(Day(ldFechaConCitas))) & "/<" & Trim(Str(lnCitasAsignadas)) & ">"
                        lnCitasAsignadas = 1
                   End If
                   ldFechaConCitas = oProgramacion.fecha
                End If
            Next
            If lnCitasAsignadas > 0 Then
               lcCitasYaAsignadas = lcCitasYaAsignadas & "/" & Trim(Str(Day(ldFechaConCitas))) & "/<" & Trim(Str(lnCitasAsignadas)) & ">"
            End If
        End If

        Set oRsTmp = Nothing
        '
        Dim lnDiaHallado As Integer, lcNroYaAsignado As String, j As Integer
        iNroDiasMes = sighentidades.diasdelmes(Year(Calendario.Value), Month(Calendario.Value))
        For i = 1 To iNroDiasMes
If i = 10 Then
lcFecha = ""
End If
            lcFecha = Trim(Str(i)) & "/" & Month(Calendario.Value) & "/" & Year(Calendario.Value)
            Calendario.DATEText(CDate(lcFecha)) = ""
            Calendario.DATEBackColor(CDate(lcFecha)) = COLOR_DIA_NO_PROGRAMADO
            CargaProgramacionDeFechaElegida False, CDate(lcFecha)
            If lcCitasYaAsignadas <> "" Then
               lnDiaHallado = InStr(lcCitasYaAsignadas, "/" & Trim(Str(i)) & "/")
               If lnDiaHallado > 0 Then
                 lcNroYaAsignado = ""
                 For j = lnDiaHallado + 1 To 1000
                     If Mid(lcCitasYaAsignadas, j, 1) = "<" Then
                        lnDiaHallado = j
                        Exit For
                     End If
                 Next
                 For j = lnDiaHallado + 1 To 1000
                     lcNroYaAsignado = lcNroYaAsignado & Mid(lcCitasYaAsignadas, j, 1)
                     If Mid(lcCitasYaAsignadas, j + 1, 1) = ">" Then
                        Exit For
                     End If
                 Next
                 
                 Calendario.DATEText(CDate(lcFecha)) = lcNroYaAsignado & "/" & Trim(Str(lnNroCuposMaximosXDia)) 'UserControl.txtCuposXdia.Text
                 Calendario.DATEBackColor(CDate(lcFecha)) = vbRed
               Else
                 If lnNroCuposMaximosXDia > 0 Then
                    Calendario.DATEText(CDate(lcFecha)) = "0/" & Trim(Str(lnNroCuposMaximosXDia))
                    Calendario.DATEBackColor(CDate(lcFecha)) = vbRed
                 End If
               End If
            Else
                 If lnNroCuposMaximosXDia > 0 Then
                    Calendario.DATEText(CDate(lcFecha)) = "0/" & Trim(Str(lnNroCuposMaximosXDia))
                    Calendario.DATEBackColor(CDate(lcFecha)) = vbRed
                 End If
            End If
        Next i
End Sub



Sub LimpiarProgramaciones()
Dim programacion As PVAppointment
Dim lKey As Long
        
        Set programacion = Diario.AppointmentSet.GetFirst()
        Do While Not programacion Is Nothing
            Calendario.DATEText(Format(programacion.StartDateTime, sighentidades.DevuelveFechaSoloFormato_DMY)) = ""
            lKey = programacion.Key
            Set programacion = Diario.AppointmentSet.GetNext(programacion)
            Diario.AppointmentSet.Remove lKey
        Loop

End Sub

Sub CargaProgramacionDeFechaElegida(lbSeCambioFechaCalendario As Boolean, ldFecha11 As Date)
    Dim oRsProgramacion As New Recordset
    mo_ReglasImagenes.SiProgramacionLlenaComboCuposMaximo lstProgramacion, mi_idSala, ldFecha11, _
                      lnNroCuposMaximosXDia, False, _
                      IIf(lbSeCambioFechaCalendario = True, sghSiProgramacionEnListBox, sghSiDevuelveTotalCuposXdia), _
                      oRsProgramacion
    Set oRsProgramacion = Nothing
End Sub

