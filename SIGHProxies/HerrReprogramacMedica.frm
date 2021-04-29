VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form HerrReprogramacMedica 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reprogramación Médica"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12225
   Icon            =   "HerrReprogramacMedica.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   12225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbNuevoIdServicioCE 
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
      Left            =   2940
      Style           =   2  'Dropdown List
      TabIndex        =   46
      Top             =   5085
      Width           =   3015
   End
   Begin VB.Frame Frame5 
      Caption         =   "Reprogramación x Paciente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6675
      Left            =   6330
      TabIndex        =   24
      Top             =   525
      Width           =   5610
      Begin VB.TextBox txtDatosDeCuenta 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   60
         TabIndex        =   35
         Top             =   2325
         Width           =   5505
      End
      Begin VB.TextBox txtNombrePaciente 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   60
         TabIndex        =   34
         Top             =   1935
         Width           =   5505
      End
      Begin VB.TextBox txtNcuenta 
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
         Left            =   1170
         MaxLength       =   30
         TabIndex        =   0
         Top             =   1470
         Width           =   1395
      End
      Begin VB.TextBox txtPlan 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   60
         TabIndex        =   33
         Top             =   2715
         Width           =   5505
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "HerrReprogramacMedica.frx":0CCA
         DownPicture     =   "HerrReprogramacMedica.frx":118E
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
         Left            =   2790
         Picture         =   "HerrReprogramacMedica.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   5880
         Width           =   1365
      End
      Begin VB.CommandButton cmdProcesaXpaciente 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "HerrReprogramacMedica.frx":1B66
         DownPicture     =   "HerrReprogramacMedica.frx":1FC6
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
         Left            =   1260
         Picture         =   "HerrReprogramacMedica.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5910
         Width           =   1365
      End
      Begin VB.ComboBox cmbIdServicioCEpac 
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
         Left            =   1155
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   690
         Width           =   4395
      End
      Begin VB.ComboBox cmbIdResponsablePac 
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
         Left            =   1155
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   315
         Width           =   4395
      End
      Begin VB.Frame Frame6 
         Caption         =   "Reprogramación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2460
         Left            =   45
         TabIndex        =   25
         Top             =   3345
         Width           =   5505
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1335
            Width           =   1245
         End
         Begin VB.ComboBox cmbIdResponsableNewPac 
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
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   435
            Width           =   3735
         End
         Begin MSMask.MaskEdBox txtFechaNewPac 
            Height          =   315
            Left            =   2340
            TabIndex        =   2
            Top             =   900
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
         Begin MSMask.MaskEdBox txtHoraFinPac 
            Height          =   315
            Left            =   4560
            TabIndex        =   38
            Top             =   1335
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
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Nueva Fecha Programada"
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
            TabIndex        =   41
            Top             =   930
            Width           =   2070
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Nuevo Médico:"
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
            TabIndex        =   40
            Top             =   465
            Width           =   1215
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "hasta"
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
            Left            =   4080
            TabIndex        =   39
            Top             =   1395
            Width           =   435
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Hora (cupo libre)"
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
            TabIndex        =   37
            Top             =   1395
            Width           =   1380
         End
      End
      Begin MSMask.MaskEdBox txtFechaAtencionPac 
         Height          =   315
         Left            =   1170
         TabIndex        =   28
         Top             =   1080
         Width           =   1395
         _ExtentX        =   2461
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "N° Cuenta"
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
         Top             =   1500
         Width           =   855
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Servicio CE"
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
         TabIndex        =   31
         Top             =   750
         Width           =   885
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "F.Atención"
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
         Top             =   1140
         Width           =   885
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
         Left            =   120
         TabIndex        =   29
         Top             =   390
         Width           =   570
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Consideraciones:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   45
      TabIndex        =   13
      Top             =   510
      Width           =   6165
      Begin VB.ListBox cmbConsideraciones 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   1530
         Left            =   90
         TabIndex        =   14
         Top             =   240
         Width           =   5940
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Programación Médica actual (con Pacientes Citados)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4650
      Left            =   15
      TabIndex        =   8
      Top             =   2550
      Width           =   6165
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "HerrReprogramacMedica.frx":28B0
         DownPicture     =   "HerrReprogramacMedica.frx":2D10
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
         Left            =   1320
         Picture         =   "HerrReprogramacMedica.frx":3185
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   3870
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "HerrReprogramacMedica.frx":35FA
         DownPicture     =   "HerrReprogramacMedica.frx":3ABE
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
         Left            =   2820
         Picture         =   "HerrReprogramacMedica.frx":3FAA
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   3870
         Width           =   1365
      End
      Begin VB.Frame Frame1 
         Caption         =   "Reprogramación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1980
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   5880
         Begin VB.CommandButton btnBuscarCEDisponible 
            Caption         =   "..."
            Height          =   360
            Left            =   2220
            TabIndex        =   47
            Top             =   705
            Width           =   525
         End
         Begin VB.ComboBox cmbIdResponsableNew 
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
            Left            =   3015
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1350
            Width           =   2805
         End
         Begin MSMask.MaskEdBox txtFechaRequeridaDesde 
            Height          =   315
            Left            =   4425
            TabIndex        =   17
            Top             =   330
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
         Begin Threed.SSOption optFecha 
            Height          =   315
            Left            =   120
            TabIndex        =   18
            Top             =   330
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   556
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Nueva Fecha Programada"
            Value           =   -1
         End
         Begin Threed.SSOption optMedico 
            Height          =   285
            Left            =   120
            TabIndex        =   19
            Top             =   1350
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   503
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Médico que reemplaza"
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Servicio CE Disponible"
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
            Left            =   375
            TabIndex        =   45
            Top             =   780
            Width           =   885
         End
      End
      Begin VB.PictureBox progressRpt 
         Height          =   300
         Left            =   4890
         ScaleHeight     =   240
         ScaleWidth      =   510
         TabIndex        =   9
         Top             =   1560
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.ComboBox cmbIdResponsable 
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
         Left            =   1155
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   315
         Width           =   4845
      End
      Begin VB.ComboBox cmbIdServicioCE 
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
         Left            =   1155
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   690
         Width           =   4830
      End
      Begin MSMask.MaskEdBox txtFechaInicio 
         Height          =   315
         Left            =   1155
         TabIndex        =   7
         Top             =   1110
         Width           =   1395
         _ExtentX        =   2461
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
      Begin MSMask.MaskEdBox txtHrInicio 
         Height          =   315
         Left            =   3540
         TabIndex        =   20
         Top             =   1110
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
      Begin MSMask.MaskEdBox txtHrFin 
         Height          =   315
         Left            =   5220
         TabIndex        =   21
         Top             =   1110
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "hasta"
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
         Left            =   4650
         TabIndex        =   23
         Top             =   1170
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Left            =   3000
         TabIndex        =   22
         Top             =   1170
         Width           =   510
      End
      Begin VB.Label Label8 
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
         Left            =   120
         TabIndex        =   12
         Top             =   390
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "F.Atención"
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
         TabIndex        =   11
         Top             =   1170
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Servicio CE"
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
         TabIndex        =   10
         Top             =   750
         Width           =   885
      End
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "En el 'Servicio CE' no deben estar registrando CITAS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   60
      TabIndex        =   42
      Top             =   30
      Width           =   11265
   End
End
Attribute VB_Name = "HerrReprogramacMedica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reprogramación médica
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasAdmision   As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_cmbIdTipoHistoria As New sighentidades.ListaDespleglable
Dim mo_cmbIdResponsable As New sighentidades.ListaDespleglable
Dim mo_cmbIdResponsableNew As New sighentidades.ListaDespleglable
Dim mo_cmbIdServicioCE As New sighentidades.ListaDespleglable
Dim mo_cmbIdResponsablePac As New sighentidades.ListaDespleglable
Dim mo_cmbIdServicioCEpac As New sighentidades.ListaDespleglable
Dim mo_cmbIdResponsableNewPac As New sighentidades.ListaDespleglable
Dim mo_clsCDOmail As New SIGHNegocios.clsCDOmail
Dim ml_idUsuario As Long
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim oRsListaMedicoNew As New Recordset
Dim oRsListaMedicoNewPac As New Recordset
Dim mo_lcNombrePc  As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lcTiempoAtencion As String
Dim lcParametro523 As String, lcParametro524 As String, lcParametro205 As String
'SCCQ 21/02/2020 Cambio7 Inicio
 Dim mo_cmbNuevoIdServicioCE As New sighentidades.ListaDespleglable
'SCCQ 21/02/2020 Cambio7 Fin
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property


Property Let idUsuario(lIdValue As Long)
    ml_idUsuario = lIdValue
End Property

Private Sub btnAceptar_Click()

If wxFranklin = "*" Then Exit Sub

    If mo_cmbIdResponsable.BoundText = "" Then
       MsgBox "Tiene que elegir al  Médico", vbInformation, "Mensaje"
       Exit Sub
    End If
    If mo_cmbIdServicioCE.BoundText = "" Then
       MsgBox "Tiene que elegir al  Servicio CE", vbInformation, "Mensaje"
       Exit Sub
    End If
    If optFecha.Value = True Then
        If txtFechaRequeridaDesde.Text = sighentidades.FECHA_VACIA_DMY Then
           MsgBox "Tiene que registrar la Nueva Fecha Programada", vbInformation, "Mensaje"
           Exit Sub
        End If
        If txtFechaRequeridaDesde.Text = txtFechaInicio.Text Then
           MsgBox "La Nueva Fecha Programada, no puede ser igual a la Fecha de Atencion", vbInformation, "Mensaje"
           Exit Sub
        End If
        If CDate(txtFechaRequeridaDesde.Text) < Date Then
           MsgBox "La Nueva Fecha Programada, no puede ser menor a la Fecha Actual", vbInformation, "Mensaje"
           Exit Sub
        End If
    Else
        If mo_cmbIdResponsableNew.BoundText = "" Then
           MsgBox "Debe elegir al Médico reemplazante", vbInformation, "Mensaje"
           Exit Sub
        End If
        If mo_cmbIdResponsableNew.BoundText = mo_cmbIdResponsable.BoundText Then
           MsgBox "El Médico reemplazante debe ser otro al Actual", vbInformation, "Mensaje"
           Exit Sub
        End If
    End If
    If MsgBox("Esta seguro", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        Me.MousePointer = 1
        On Error GoTo ErrorProceso
        Dim oRsTmp As New Recordset
        Dim oRsTmp1 As New Recordset
        Dim oRsTmp2 As New Recordset
        Dim oConexion As New Connection
        Dim oConexionExterna As New Connection
        Dim lbPasaAtencion As Boolean
        Dim lnAtencionesPasadas As Integer
        Dim lnAtencionesAtendidos  As Integer
        Dim lnAtencionesFarmaciaServicios As Integer
        Dim lnIdProgramacion As Long
        Dim lnIdProgramacionNew As Long
        Dim oDOProgramacionMedica As New DOProgramacionMedica
        Dim oProgramacionMedica As New ProgramacionMedica
        Dim lbHuboCitadoFueraDeHora As Boolean
        Dim lcSql As String
        'SCCQ 21/02/2020 Cambio7 Inicio
        Dim lnIdServicioCENew As Long
        'SCCQ 21/02/2020 Cambio7 Fin
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        oConexionExterna.CommandTimeout = 300
        oConexionExterna.CursorLocation = adUseClient
        oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
        
        
        'debb-12/04/2016 (inicio)
        If optMedico.Value = True Then
            Set oRsTmp = mo_ReglasDeProgMedica.ProgramacionMedicaSeleccionarPorMedicoFechaServicio(Val(mo_cmbIdResponsableNew.BoundText), txtFechaInicio.Text, 0)
            If oRsTmp.RecordCount > 0 Then
                oRsTmp.MoveFirst
                Do While Not oRsTmp.EOF
                   If (Me.txtHrInicio.Text >= oRsTmp!HoraInicio And Me.txtHrInicio.Text <= oRsTmp!HoraFin) Or _
                      (Me.txtHrFin.Text >= oRsTmp!HoraInicio And Me.txtHrFin.Text <= oRsTmp!HoraFin) Then
                        MsgBox "Yá se programo al 'Médico que reemplaza' ese mismo día, turno", vbInformation, "Mensaje"
                        oConexion.Close
                        oConexionExterna.Close
                        Exit Sub
                   End If
                oRsTmp.MoveNext
                Loop
            End If
            oRsTmp.Close
        Else
            'debb-19/09/2019
            'SCCQ 21/02/2020 Cambio7 Inicio
            If Val(mo_cmbNuevoIdServicioCE.BoundText) = 0 Then 'No seleccionó ningún Servicio de CE disponible
                MsgBox "Debe elegir un consultorio disponible", vbInformation, "Mensaje"
                Exit Sub
            End If
            lnIdServicioCENew = Val(mo_cmbNuevoIdServicioCE.BoundText)
           
            Set oRsTmp = mo_ReglasDeProgMedica.ProgramacionMedicaSeleccionarXfechaConsultorio(txtFechaRequeridaDesde.Text, lnIdServicioCENew)
             'SCCQ 21/02/2020 Cambio7 Fin
            If oRsTmp.RecordCount > 0 Then
                oRsTmp.MoveFirst
                Do While Not oRsTmp.EOF
                   If (Me.txtHrInicio.Text >= oRsTmp!HoraInicio And Me.txtHrInicio.Text <= oRsTmp!HoraFin) Or _
                      (Me.txtHrFin.Text >= oRsTmp!HoraInicio And Me.txtHrFin.Text <= oRsTmp!HoraFin) Then
                          If oRsTmp!idMedico <> Val(mo_cmbIdResponsable.BoundText) Then
                                MsgBox "Yá se programo a otro 'Médico' ese mismo día, hora inicio", vbInformation, "Mensaje"
                                oConexion.Close
                                oConexionExterna.Close
                                'SCCQ 20/02/2020 Cambio7 Inicio
                                'BORRAR DATOS DE cmbNuevoIdServicioCE
                                cmbNuevoIdServicioCE.Clear
                                'SCCQ 20/02/2020 Cambio7 Fin
                                Exit Sub
                          End If
                   End If
                   oRsTmp.MoveNext
                Loop
            End If
            oRsTmp.Close
            'debb-19/09/2019
            Set oRsTmp = mo_ReglasDeProgMedica.ProgramacionMedicaSeleccionarPorMedicoFechaServicio(Val(mo_cmbIdResponsable.BoundText), txtFechaRequeridaDesde.Text, 0)
            If oRsTmp.RecordCount > 0 Then
                oRsTmp.MoveFirst
                Do While Not oRsTmp.EOF
                   If (Me.txtHrInicio.Text >= oRsTmp!HoraInicio And Me.txtHrInicio.Text <= oRsTmp!HoraFin) Or _
                      (Me.txtHrFin.Text >= oRsTmp!HoraInicio And Me.txtHrFin.Text <= oRsTmp!HoraFin) Then
                        MsgBox "Yá se programo al 'Médico' ese mismo día", vbInformation, "Mensaje"
                        oConexion.Close
                        'SCCQ 26-04-2021 Cambio 64 Inicio (Reversion Cambio 46)
                        oConexionExterna.CloDoEvents
                        'RHA 20/01/2021 Cambio 46 Inicio
                         'Antes:oConexionExterna.CloDoEvents
                        'oConexionExterna.Close
                        'RHA 18/01/2021 Cambio 46 Fin
                        'SCCQ 26-04-2021 Cambio 64 Fin (Reversion Cambio 46)
                        
                        'SCCQ 20/02/2020 Cambio7 Inicio
                        'BORRAR DATOS DE cmbNuevoIdServicioCE
                        cmbNuevoIdServicioCE.Clear
                        'SCCQ 20/02/2020 Cambio7 Fin
                        Exit Sub
                   End If
                   oRsTmp.MoveNext
                Loop
            End If
            oRsTmp.Close
        End If
        'debb-12/04/2016 (fin)
        '
        Set oRsTmp = mo_ReglasDeProgMedica.ProgramacionMedicaSeleccionarPorMedicoFechaServicio(Val(mo_cmbIdResponsable.BoundText), txtFechaInicio.Text, Val(mo_cmbIdServicioCE.BoundText))
        oRsTmp.Filter = "horaInicio='" & Me.txtHrInicio.Text & "'"
        If oRsTmp.RecordCount = 0 Then
           MsgBox "Esa Programación Médica NO EXISTE", vbInformation, "Mensaje"
           oConexion.Close
           oConexionExterna.Close
           'SCCQ 20/02/2020 Cambio7 Inicio
           'BORRAR DATOS DE cmbNuevoIdServicioCE
           cmbNuevoIdServicioCE.Clear
           'SCCQ 20/02/2020 Cambio7 Fin
           Exit Sub
        End If
        lnIdProgramacion = oRsTmp.Fields!idProgramacion
        oRsTmp.Close
        '
        lbHuboCitadoFueraDeHora = False
        oConexionExterna.BeginTrans
        oConexion.BeginTrans
        Set oProgramacionMedica.Conexion = oConexion
        If optFecha.Value = True Then
            '********Por Fecha
            Set oRsTmp1 = mo_ReglasAdmision.AtencionesCEseleccionarPorFechaServicioMedico(txtFechaInicio.Text, Val(mo_cmbIdServicioCE.BoundText), Val(mo_cmbIdResponsable.BoundText))
            lnAtencionesPasadas = 0
            lnAtencionesAtendidos = 0
            lnIdProgramacionNew = 0
            lnAtencionesFarmaciaServicios = 0
            If oRsTmp1.RecordCount > 0 Then
               oRsTmp1.MoveFirst
               Do While Not oRsTmp1.EOF
                    lbPasaAtencion = False
                    If IsNull(oRsTmp1.Fields!FechaEgreso) Then
                        Set oRsTmp2 = mo_ReglasFarmacia.FarmMovimientoVentasSeleccionarPorCuenta(oRsTmp1.Fields!idCuentaAtencion)
                        If oRsTmp2.RecordCount = 0 Then
                           oRsTmp2.Close
                           Set oRsTmp2 = mo_ReglasFacturacion.FactOrdenServicioSeleccionarPorIdCuenta(oRsTmp1.Fields!idCuentaAtencion)
                           oRsTmp2.Filter = "idEstadoFacturacion<>9"
                          'SCCQ 26-04-2021 Cambio 64 Inicio (Reversion Cambio 46)
                          If oRsTmp2.RecordCount <= 1 Then
                          'RHA 20/01/2021 Cambio 46 Inicio
                         'Antes: If oRsTmp2.RecordCount <= 1  Then
                           'If oRsTmp2.RecordCount = 3 Or oRsTmp2.RecordCount = 1 Then
                         'RHA 20/01/2021 Cambio 46 Fin
                          'SCCQ 26-04-2021 Cambio 64 Fin (Reversion Cambio 46)
                              lbPasaAtencion = True
                           Else
                              lnAtencionesFarmaciaServicios = lnAtencionesFarmaciaServicios + 1
                              lbHuboCitadoFueraDeHora = True
                           End If
                        Else
                           lnAtencionesFarmaciaServicios = lnAtencionesFarmaciaServicios + 1
                           lbHuboCitadoFueraDeHora = True
                        End If
                        oRsTmp2.Close
                        If lbPasaAtencion = True Then
                            If (oRsTmp1.Fields!HoraIngreso >= txtHrInicio.Text And oRsTmp1.Fields!HoraIngreso <= txtHrFin.Text) Then
                            Else
                               lbPasaAtencion = False
                               lbHuboCitadoFueraDeHora = True
                            End If
                        End If
                    End If
                    If lbPasaAtencion = True Then
                       If lnAtencionesPasadas = 0 Then
                            oDOProgramacionMedica.idProgramacion = lnIdProgramacion
                            If oProgramacionMedica.SeleccionarPorId(oDOProgramacionMedica) Then
                               oDOProgramacionMedica.fecha = txtFechaRequeridaDesde.Text
                               'SCCQ 21/02/2020 Cambio7 Inicio
                               'Inserta nueva programación, se ingresa el nuevo consultorio disponible mo_cmbNuevoIdServicioCE
                               'Asignamos el valor del nuevo idServicioCEDisponible
                               oDOProgramacionMedica.idServicio = Val(mo_cmbNuevoIdServicioCE.BoundText)
                               'SCCQ 21/02/2020 Cambio7 Fin
                               If oProgramacionMedica.Insertar(oDOProgramacionMedica) Then
                                  lnIdProgramacionNew = oDOProgramacionMedica.idProgramacion
                               Else
                                  MsgBox oProgramacionMedica.MensajeError
                                  GoTo ErrorProceso
                               End If
                            Else
                               MsgBox oProgramacionMedica.MensajeError
                               GoTo ErrorProceso
                            End If
                       End If
                       lnAtencionesPasadas = lnAtencionesPasadas + 1
                      'SCCQ 21/02/2020 Cambio7 Inicio
                       'Actualiza citas con nueva programación, pasar nuevos datos de idservicioCEdisponible que debe ser igual a la tabla progrmacion
                       lnIdServicioCENew = Val(mo_cmbNuevoIdServicioCE.BoundText)
                       mo_ReglasDeProgMedica.CitasActualizaDatosDeReprogramacionXfechaServicioCE txtFechaRequeridaDesde.Text, _
                                                             oRsTmp1.Fields!idAtencion, lnIdProgramacionNew, lnIdServicioCENew, oConexion
                       
                       'SCCQ 21/02/2020 Cambio7 Fin
                      ' If oRsTmp1!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                         ' ActualizaMedicoEnFuasYaEmitidas oConexionExterna, oRsTmp1!idCuentaAtencion, _
                                                       0, txtFechaRequeridaDesde.Text, ""
                      ' End If
                       
                       EnviaEmail oConexion, oRsTmp1!idPaciente, txtFechaRequeridaDesde.Text & " " & oRsTmp1.Fields!HoraIngreso, _
                                  txtFechaInicio.Text, oRsTmp1!idCuentaAtencion
                    Else
                       lnAtencionesAtendidos = lnAtencionesAtendidos + 1
                    End If
                    oRsTmp1.MoveNext
               Loop
            End If
            If lnAtencionesPasadas = 0 Then
               MsgBox "No se pudo Reprogramar" & Chr(13) & "Pacientes Citados: " & Trim(str(oRsTmp1.RecordCount)) & Chr(13) & "Pacientes Atendidos: " & Trim(str(lnAtencionesAtendidos)), vbInformation, "Mensaje"
               oConexion.RollbackTrans
               oConexion.Close
               Exit Sub
            Else
               If lbHuboCitadoFueraDeHora = False Then
                  mo_ReglasDeProgMedica.ProgramacionMedicaEliminarPorId lnIdProgramacion, oConexion
               End If
               '
               Dim mo_Procesos As New SIGHProxies.Procesos
               Dim lcMensaje As String, lbSeTerminaSistema As Boolean, oRsCitasWeb As New Recordset, lcUsuario As String
               lcMensaje = ""
               Set mo_Procesos = Nothing
               '
            End If
            oRsTmp1.Close
            
        Else
            '********Por otro Medico
            Set oRsTmp1 = mo_ReglasAdmision.AtencionesCEseleccionarPorFechaServicioMedico(txtFechaInicio.Text, Val(mo_cmbIdServicioCE.BoundText), Val(mo_cmbIdResponsable.BoundText))
            oRsTmp1.Filter = "horaIngreso>='" & Me.txtHrInicio.Text & "' and horaIngreso<='" & Me.txtHrFin.Text & "'"

            lnAtencionesPasadas = 0
            lnAtencionesAtendidos = 0
            lnIdProgramacionNew = 0
            lnAtencionesFarmaciaServicios = 0
            If oRsTmp1.RecordCount > 0 Then
               oRsTmp1.MoveFirst
               Do While Not oRsTmp1.EOF
                    lbPasaAtencion = False
                    If IsNull(oRsTmp1.Fields!FechaEgreso) Then
                        Set oRsTmp2 = mo_ReglasFarmacia.FarmMovimientoVentasSeleccionarPorCuenta(oRsTmp1.Fields!idCuentaAtencion)
                        If oRsTmp2.RecordCount = 0 Then
                           oRsTmp2.Close
                           Set oRsTmp2 = mo_ReglasFacturacion.FactOrdenServicioSeleccionarPorIdCuenta(oRsTmp1.Fields!idCuentaAtencion)
                           oRsTmp2.Filter = "idEstadoFacturacion<>9"
                         'SCCQ 26-04-2021 Cambio 64 Inicio (Reversion Cambio 46)
                         If oRsTmp2.RecordCount <= 1 Then
                         'RHA 20/01/2021 Cambio 46 Inicio
                         'Antes: If oRsTmp2.RecordCount <= 1  Then
                           'If oRsTmp2.RecordCount = 3 Or oRsTmp2.RecordCount = 1 Then
                         'RHA 20/01/2021 Cambio 46 Fin
                         'SCCQ 26-04-2021 Cambio 64 Fin (Reversion Cambio 46)
                              lbPasaAtencion = True
                           Else
                              lnAtencionesFarmaciaServicios = lnAtencionesFarmaciaServicios + 1
                              'SCCQ 26-04-2021 Cambio 64 Inicio (Reversion Cambio 46)
                              lbHuboCitadoFueraDeHora = True 'Código sin firma
                              'SCCQ 26-04-2021 Cambio 64 Fin (Reversion Cambio 46)
                           End If
                        Else
                           lnAtencionesFarmaciaServicios = lnAtencionesFarmaciaServicios + 1
                           'SCCQ 26-04-2021 Cambio 64 Inicio (Reversion Cambio 46)
                           lbHuboCitadoFueraDeHora = True 'Código sin firma
                           'SCCQ 26-04-2021 Cambio 64 Fin (Reversion Cambio 46)
                        End If
                        oRsTmp2.Close
                        If lbPasaAtencion = True Then
                            If (oRsTmp1.Fields!HoraIngreso >= txtHrInicio.Text And oRsTmp1.Fields!HoraIngreso <= txtHrFin.Text) Then
                            Else
                               lbPasaAtencion = False
                               lbHuboCitadoFueraDeHora = True
                            End If
                        End If
                    End If
                    If lbPasaAtencion = True Then
                       If lnAtencionesPasadas = 0 Then
                            oDOProgramacionMedica.idProgramacion = lnIdProgramacion
                            If oProgramacionMedica.SeleccionarPorId(oDOProgramacionMedica) Then
                               oDOProgramacionMedica.idMedico = Val(mo_cmbIdResponsableNew.BoundText)
                               If oProgramacionMedica.Insertar(oDOProgramacionMedica) Then
                                  lnIdProgramacionNew = oDOProgramacionMedica.idProgramacion
                               Else
                                  MsgBox oProgramacionMedica.MensajeError
                                  GoTo ErrorProceso
                               End If
                            Else
                               MsgBox oProgramacionMedica.MensajeError
                               GoTo ErrorProceso
                            End If
                       End If
                       lnAtencionesPasadas = lnAtencionesPasadas + 1
                       mo_ReglasDeProgMedica.CitasActualizaDatosDeReporgramacionXmedico Val(mo_cmbIdResponsableNew.BoundText), _
                                                                     lnIdProgramacionNew, oRsTmp1.Fields!idAtencion, oConexion
                       
                       If oRsTmp1!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                          ActualizaMedicoEnFuasYaEmitidas oConexionExterna, oRsTmp1!idCuentaAtencion, _
                                                       Val(mo_cmbIdResponsableNew.BoundText), 0, ""
                       End If
                    Else
                       lnAtencionesAtendidos = lnAtencionesAtendidos + 1
                    End If
                    oRsTmp1.MoveNext
               Loop
            End If
            If lnAtencionesPasadas = 0 Then
               MsgBox "No se pudo Reprogramar" & Chr(13) & "Pacientes Citados: " & Trim(str(oRsTmp1.RecordCount)) & Chr(13) & "Pacientes Atendidos: " & Trim(str(lnAtencionesAtendidos)), vbInformation, "Mensaje"
               oConexion.RollbackTrans
               oConexion.Close
               Exit Sub
            Else
               If lbHuboCitadoFueraDeHora = False Then
                  mo_ReglasDeProgMedica.ProgramacionMedicaEliminarPorId lnIdProgramacion, oConexion
               End If
            End If
            oRsTmp1.Close
        End If
        '******************************Actualiza Cupos Web (inicio)**********************************************
        Dim oDOCitasWebCupos As New DOCitasWebCupos
        Dim oCitasWebCupos As New CitasWebCupos
        Dim oDOCitaBloqueada As New DOCitaBloqueada
        Dim oCitasBloqueadas As New CitasBloqueadas
        Dim lcMensajeWeb As String
        Set oCitasBloqueadas.Conexion = oConexion
        Set oCitasWebCupos.Conexion = oConexionExterna
        Set oRsTmp1 = mo_ReglasDeProgMedica.CitasWebCuposSeleccionarPorFechas(CDate(txtFechaInicio.Text), _
                                           CDate(txtFechaInicio.Text), Val(mo_cmbIdResponsable.BoundText), _
                                           Val(mo_cmbIdServicioCE.BoundText), oConexionExterna)
      
         lcMensajeWeb = ""
         If oRsTmp1.RecordCount > 0 Then
            oRsTmp1.MoveFirst
            Do While Not oRsTmp1.EOF
               If (oRsTmp1.Fields!HoraInicio >= txtHrInicio.Text And oRsTmp1.Fields!HoraFinal <= txtHrFin.Text) And Not IsNull(oRsTmp1!idCitaBloqueada) Then
                    oDOCitaBloqueada.idCitaBloqueada = oRsTmp1!idCitaBloqueada
                    If oCitasBloqueadas.SeleccionarPorId(oDOCitaBloqueada) = True Then
                       oDOCitaBloqueada.IdUsuarioAuditoria = sighentidades.Usuario
                       If optFecha.Value = True Then
                          oDOCitaBloqueada.fecha = CDate(txtFechaRequeridaDesde.Text)
                       Else
                          oDOCitaBloqueada.idMedico = Val(mo_cmbIdResponsableNew.BoundText)
                       End If
                       If oCitasBloqueadas.Modificar(oDOCitaBloqueada) = True Then
                       End If
                    End If
                    Set oDOCitasWebCupos = oCitasWebCupos.SeleccionarPorIdCitaBloqueada(oRsTmp1!idCitaBloqueada)
                    If Not (oDOCitasWebCupos Is Nothing) Then
                       oDOCitasWebCupos.IdUsuarioAuditoria = sighentidades.Usuario
                       If optFecha.Value = True Then
                          oDOCitasWebCupos.fecha = CDate(txtFechaRequeridaDesde.Text)
                          lcMensajeWeb = "Hubo cambio de FECHA en CUPOS WEB" & Chr(13) & _
                                       "se sugiere volver a en enviar los CUPOS a la WEB" & Chr(13) & _
                                       "en Herramientas -> Citas Web Configurar -> pestaña EXPORTAR DATOS"
                       Else
                          oDOCitasWebCupos.idMedico = Val(mo_cmbIdResponsableNew.BoundText)
                          lcMensajeWeb = "Hubo cambio de MEDICO en CUPOS WEB" & Chr(13) & _
                                       "se sugiere volver a en enviar los CUPOS a la WEB" & Chr(13) & _
                                       "en Herramientas -> Citas Web Configurar -> pestaña EXPORTAR DATOS"
                       End If
                       If oCitasWebCupos.Modificar(oDOCitasWebCupos) = True Then
                       End If
                    End If
               End If
               oRsTmp1.MoveNext
            Loop
         End If
         If lcMensajeWeb <> "" Then
            MsgBox lcMensajeWeb, vbInformation, ""
         End If
         Set oDOCitasWebCupos = Nothing
         Set oCitasWebCupos = Nothing
         Set oDOCitaBloqueada = Nothing
         Set oCitasBloqueadas = Nothing
         '******************************Actualiza Cupos Web (fin)**********************************************
         
         '
         Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
         Call mo_ReglasSeguridad.AuditoriaAgregarV(ml_idUsuario, "M", lnIdProgramacion, "ProgramacionMedica", oConexion, 500 + 128, mo_lcNombrePc, "Reprog.Med: " & Trim(cmbIdResponsable.Text) & "  Serv: " & Trim(cmbIdServicioCE.Text))     '500+ ListBarReporte.idReporte
         oConexion.CommitTrans
         oConexion.Close
         oConexionExterna.CommitTrans
         oConexionExterna.Close
         Me.MousePointer = 11
         Me.Visible = False
    End If
    Set oRsTmp = Nothing
    Set oRsTmp1 = Nothing
    Set oRsTmp2 = Nothing
    Set oConexion = Nothing
    Set oConexionExterna = Nothing
    Exit Sub
ErrorProceso:
    oConexion.RollbackTrans
    oConexionExterna.RollbackTrans
    If lnAtencionesFarmaciaServicios > 0 Then
       MsgBox "No se puede REPROGRAMAR, ya que hubo Atenciones en Farmacia/Imagenes/Laboratorio para algunos Pacientes: " & Trim(str(lnAtencionesFarmaciaServicios)), vbInformation, Me.Caption
    End If
    MsgBox Err.Description
    Exit Sub
    Resume

End Sub

'SCCQ 19/02/2020 Cambio 7 Inicio
Private Sub btnBuscarCEDisponible_Click()
LlenarCmbNuevoIdServicioCE
End Sub
Private Sub LlenarCmbNuevoIdServicioCE()
If mo_cmbIdServicioCE.BoundText <> "" Then
    Dim mo_AdminServHosp As New ReglasServiciosHosp
    Set mo_cmbNuevoIdServicioCE.MiComboBox = cmbNuevoIdServicioCE
       mo_cmbNuevoIdServicioCE.BoundColumn = "idServicio"
       mo_cmbNuevoIdServicioCE.ListField = "descripcionLarga"
        'Seleccionar el idEspecidad del ServicioCE seleccionado INICIO
        Dim oDOServicio As New doServicio
        Dim oConexion As New Connection
        oConexion.Open sighentidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        Set oDOServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(Val(mo_cmbIdServicioCE.BoundText), oConexion)
        oConexion.Close
        Set oConexion = Nothing
        'Seleccionar el idEspecidad del ServicioCE seleccionado FIN
       Set mo_cmbNuevoIdServicioCE.RowSource = mo_AdminServHosp.ServiciosSeleccionarCEDisponibles(oDOServicio.IdEspecialidad, txtHrInicio.Text, txtHrFin.Text, txtFechaRequeridaDesde.Text)
        If cmbNuevoIdServicioCE.ListCount > 0 Then
            cmbNuevoIdServicioCE.ListIndex = 0
        Else
            MsgBox "No hay consultorios disponibles para esa fecha", vbInformation, "Mensaje"
       End If
Else
 MsgBox "Seleccione Servicio CE", vbInformation, "Mensaje"
End If
End Sub
'SCCQ 19/02/2020 Cambio 7 Fin
Private Sub btnCancelar_Click()
     Me.Visible = False
End Sub

Private Sub cmbIdResponsableNewPac_KeyDown(KeyCode As Integer, Shift As Integer)
        mo_Teclado.RealizarNavegacion KeyCode, cmbIdResponsableNewPac
End Sub

Private Sub cmbIdResponsableNewPac_LostFocus()
    LlenaComboConHoraInicio
End Sub

Private Sub cmbIdServicioCE_Click()
   If mo_cmbIdServicioCE.BoundText <> "" Then
       Dim oBuscaMedicos As New SIGHNegocios.ReglasDeProgMedica
       Dim lnIdEspecialidad As Long
       Dim lcSql As String
       If oRsListaMedicoNew.State = 1 Then
          oRsListaMedicoNew.Close
       End If
       Set oRsListaMedicoNew = mo_ReglasFacturacion.ServiciosSeleccionarPorFiltro("idServicio=" & mo_cmbIdServicioCE.BoundText, sghPorCodigo)
       lnIdEspecialidad = 0
       If oRsListaMedicoNew.RecordCount > 0 Then
          lnIdEspecialidad = oRsListaMedicoNew.Fields!IdEspecialidad
       End If
       oRsListaMedicoNew.Close
       Set oRsListaMedicoNew = mo_ReglasDeProgMedica.MedicosSeleccionarPorEspecialidad(lnIdEspecialidad)
       oRsListaMedicoNew.Filter = "esActivo=true"
       mo_cmbIdResponsableNew.BoundColumn = "IdMedico"
       mo_cmbIdResponsableNew.ListField = "Dmedico"
       Set mo_cmbIdResponsableNew.RowSource = oRsListaMedicoNew
  End If
End Sub





Private Sub cmbIdServicioCEpac_Click()
   If mo_cmbIdServicioCEpac.BoundText <> "" Then
       Dim oBuscaMedicos As New SIGHNegocios.ReglasDeProgMedica
       Dim lnIdEspecialidad As Long
       Dim lcSql As String
       If oRsListaMedicoNewPac.State = 1 Then
          oRsListaMedicoNewPac.Close
       End If
       Set oRsListaMedicoNewPac = mo_ReglasFacturacion.ServiciosSeleccionarPorFiltro("idServicio=" & mo_cmbIdServicioCEpac.BoundText, sghPorCodigo)
       lnIdEspecialidad = 0
       If oRsListaMedicoNewPac.RecordCount > 0 Then
          lnIdEspecialidad = oRsListaMedicoNewPac.Fields!IdEspecialidad
       End If
       oRsListaMedicoNewPac.Close
       Set oRsListaMedicoNewPac = mo_ReglasDeProgMedica.MedicosSeleccionarPorEspecialidad(lnIdEspecialidad)
       oRsListaMedicoNewPac.Filter = "esActivo=true"
       mo_cmbIdResponsableNewPac.BoundColumn = "IdMedico"
       mo_cmbIdResponsableNewPac.ListField = "Dmedico"
       Set mo_cmbIdResponsableNewPac.RowSource = oRsListaMedicoNewPac
       '
       lcTiempoAtencion = mo_AdminServiciosHosp.EspecialidadCEseleccionarIdServicio(Val(mo_cmbIdServicioCEpac.BoundText))
  End If

End Sub

Private Sub cmdCancelar_Click()
     Me.Visible = False
End Sub

Private Sub cmdProcesaXpaciente_Click()

If wxFranklin = "*" Then Exit Sub

    On Error GoTo ErrorProceso1
    Dim oRsCita As New Recordset
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim oRsTmp3 As New Recordset
    Dim rsCitas As New Recordset
    Dim oDoCita As New doCita, oCitas As New Citas
    Dim oConexion As New Connection
    Dim oConexionExterna As New Connection
    Dim oReglasArchivoClinico As New ReglasArchivoClinico
    Dim lcSql As String
    Dim lbPasaAtencion As Boolean
    Dim lnAtencionesPasadas As Integer
    Dim lnAtencionesAtendidos  As Integer
    Dim lnAtencionesFarmaciaServicios As Integer
    Dim lnIdProgramacion As Long
    Dim lnIdProgramacionNew As Long
    Dim lbHuboCitadoFueraDeHora As Boolean
    Dim lnIdAtencion As Long, lnIdEspecialidad As Long
    Dim lnIdServicioNew As Long, lnIdProductoNew As Long
    Dim lbEsCitaPagada As Boolean
   
    
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    oConexionExterna.CommandTimeout = 300
    oConexionExterna.CursorLocation = adUseClient
    oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    
    
    
    If txtPlan.Text = "" Then
       MsgBox "Tiene que ingresar el NRO DE CUENTA", vbInformation, "Mensaje"
       Exit Sub
    End If
    If mo_cmbIdResponsablePac.BoundText = "" Then
       MsgBox "Tiene que elegir al  Médico", vbInformation, "Mensaje"
       Exit Sub
    End If
    If mo_cmbIdServicioCEpac.BoundText = "" Then
       MsgBox "Tiene que elegir al  Servicio CE", vbInformation, "Mensaje"
       Exit Sub
    End If
    If txtFechaNewPac.Text = sighentidades.FECHA_VACIA_DMY Then
       MsgBox "Tiene que registrar la Nueva Fecha Programada", vbInformation, "Mensaje"
       Exit Sub
    End If
    If txtHoraPac.Text = sighentidades.HORA_VACIA_HM Then
       MsgBox "Tiene que registrar la Nueva Hora Programada", vbInformation, "Mensaje"
       Exit Sub
    End If
    'If CDate(txtFechaAtencionPac.Text) < Date Then
    If CDate(txtFechaNewPac.Text) < Date Then
       MsgBox "La Nueva Fecha Programada, no puede ser menor a la Fecha Actual", vbInformation, "Mensaje"
       Exit Sub
    End If
    If mo_cmbIdResponsableNewPac.BoundText = "" Then
       MsgBox "Debe elegir al Médico reemplazante", vbInformation, "Mensaje"
       Exit Sub
    End If
    Set oRsCita = mo_ReglasDeProgMedica.CitasSeleccionarXfechaMedico(Me.txtFechaNewPac.Text, Val(mo_cmbIdResponsableNewPac.BoundText))
    If oRsCita.RecordCount > 0 Then
       lnIdEspecialidad = oRsCita.Fields!IdEspecialidad
       lnIdServicioNew = oRsCita.Fields!idServicio
       lnIdProgramacionNew = oRsCita.Fields!idProgramacion
       lnIdProductoNew = oRsCita.Fields!idProducto
       lcSql = "HoraInicio='" & Me.txtHoraPac.Text & "'"
       oRsCita.Find lcSql
       If Not oRsCita.EOF Then
          Do While Not oRsCita.EOF
            'If oRsCita.Fields!HoraInicio = Me.txtHoraPac.Text And oRsCita.Fields!HoraFin = txtHoraFinPac.Text Then
            If oRsCita.Fields!HoraInicio >= Me.txtHoraPac.Text And oRsCita.Fields!HoraFin < Me.txtHoraPac.Text Then
                oRsCita.Close
                MsgBox "Ya se programó esa Fecha para el NUEVO MEDICO, pero la Hora se está ocupada", vbInformation, "Mensaje"
                Exit Sub
            End If
            oRsCita.MoveNext
          Loop
          
       End If
       'If cmbIdResponsablePac.Text = cmbIdResponsableNewPac And txtFechaAtencionPac = txtFechaNewPac Then
            oRsCita.Close
            Set oRsCita = mo_ReglasDeProgMedica.ProgramacionMedicaSeleccionarXFechaMedico(Me.txtFechaNewPac.Text, Val(mo_cmbIdResponsableNewPac.BoundText))
            If oRsCita.RecordCount > 0 Then
               oRsCita.MoveFirst
               Do While Not oRsCita.EOF
                  If Me.txtHoraPac.Text >= oRsCita!HoraInicio And Me.txtHoraPac.Text <= oRsCita!HoraFin Then
                     lnIdProgramacionNew = oRsCita.Fields!idProgramacion
                     Exit Do
                  End If
                  oRsCita.MoveNext
               Loop
            End If
       'End If
    Else
       
       oRsCita.Close
       Set oRsCita = mo_ReglasDeProgMedica.ProgramacionMedicaSeleccionarXFechaMedico(Me.txtFechaNewPac.Text, Val(mo_cmbIdResponsableNewPac.BoundText))
       If oRsCita.RecordCount > 0 Then
            lnIdEspecialidad = oRsCita.Fields!IdEspecialidadMedico
            lnIdServicioNew = oRsCita.Fields!idServicio
            lnIdProgramacionNew = oRsCita.Fields!idProgramacion
            lnIdProductoNew = oRsCita.Fields!IdProductoConsulta   '4583
       Else
            MsgBox "No se programó esa Fecha para el NUEVO Médico", vbInformation, "Mensaje"
            Exit Sub
       End If
    End If
    Set oRsTmp1 = mo_ReglasAdmision.AtencionesCEseleccionarPorFechaServicioMedicoCuenta(Me.txtFechaAtencionPac.Text, Val(mo_cmbIdServicioCEpac.BoundText), Val(mo_cmbIdResponsablePac.BoundText), Val(Me.txtNcuenta.Text))
    If oRsTmp1.RecordCount = 0 Then
        oRsTmp1.Close
        MsgBox "Ese Paciente no tiene CITA actual", vbInformation, "Mensaje"
        Exit Sub
    End If
    If IsNull(oRsTmp1.Fields!FechaEgreso) Then
        Set oRsTmp2 = mo_ReglasFarmacia.FarmMovimientoVentasSeleccionarPorCuenta(Val(Me.txtNcuenta.Text))
        If oRsTmp2.RecordCount = 0 Then
           oRsTmp2.Close
           Set oRsTmp2 = mo_ReglasFacturacion.FactOrdenServicioSeleccionarPorIdCuenta(Val(Me.txtNcuenta.Text))
           oRsTmp2.Filter = "idEstadoFacturacion<>9"
           'RHA 18/01/2021 Cambio46 Inicio
           'Antes:If oRsTmp2.RecordCount <= 1 Then
           If oRsTmp2.RecordCount <= 3 Then
           'RHA 18/01/2021 Cambio46 Fin
           Else
                If sighentidades.Parametro561 = "S" Then
                    oRsTmp2.Close
                    Set oRsTmp2 = mo_ReglasAdmision.AtencionesDiagnosticosSeleccionarPorNroCuenta(Val(Me.txtNcuenta.Text))
                    If oRsTmp2.RecordCount > 0 Then
                        MsgBox "Ese paciente tiene registrado CPT y ya le registraron DX (ya lo atendieron en CE)", vbInformation, Me.Caption
                        Exit Sub
                    End If
                Else
                    MsgBox "Ese paciente tubo despachos en Procedimientos", vbInformation, Me.Caption
                    Exit Sub
                End If
           End If
        Else
            MsgBox "Ese paciente tubo despachos en Farmacia", vbInformation, Me.Caption
            Exit Sub
        End If
        oRsTmp2.Close
    Else
        MsgBox "Ese paciente ya fué atendido por el MEDICO", vbInformation, Me.Caption
        Exit Sub
    End If
    If MsgBox("Esta seguro", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        'verifica que otro Paciente ya ocupe la HORA de la CITA
        Set rsCitas = mo_ReglasAdmision.PacienteTieneCita(CDate(txtFechaNewPac.Text), Val(mo_cmbIdServicioCEpac.BoundText), 0)
        If rsCitas.RecordCount > 0 Then
                rsCitas.MoveFirst
                rsCitas.Find "horaInicio='" & txtHoraPac.Text & "'"
                If Not rsCitas.EOF Then
                     MsgBox "Ya existe CITA para otro paciente en esa Hora: " & txtHoraPac.Text, vbInformation, Me.Caption
                     Exit Sub
                End If
        End If
        rsCitas.Close
        '
        Me.MousePointer = 1
        '
        lbHuboCitadoFueraDeHora = False
        If oRsTmp1.RecordCount > 0 Then
           '
           lbEsCitaPagada = False
           If oRsTmp1!IdFormaPago = 1 Then
                Set oRsTmp3 = mo_AdminCaja.CajaComprobantesPagoXcuenta(oRsTmp1!idCuentaAtencion, oConexion)
                If oRsTmp3.RecordCount > 0 Then
                   If oRsTmp3!IdEstadoComprobante = 4 Then
                      lbEsCitaPagada = True
                   End If
                End If
                oRsTmp3.Close
           End If
           
           '
           oConexion.BeginTrans
           oDoCita.fecha = Me.txtFechaNewPac.Text
           oDoCita.HoraInicio = Me.txtHoraPac.Text
           oDoCita.HoraFin = Me.txtHoraFinPac.Text
           oDoCita.idPaciente = oRsTmp1.Fields!idPaciente
           oDoCita.IdEstadoCita = IIf(lbEsCitaPagada = True, 4, 1)
           oDoCita.idAtencion = oRsTmp1.Fields!idAtencion
           oDoCita.idMedico = Val(mo_cmbIdResponsableNewPac.BoundText)
           oDoCita.IdEspecialidad = lnIdEspecialidad      'oRsCita.Fields!IdEspecialidad
           oDoCita.idServicio = lnIdServicioNew           'oRsCita.Fields!idServicio
           oDoCita.idProgramacion = lnIdProgramacionNew   'oRsCita.Fields!IdProgramacion
           oDoCita.idProducto = lnIdProductoNew           'oRsCita.Fields!idProducto
           oDoCita.FechaSolicitud = Format(Date, "dd/mm/yyyy")
           oDoCita.HoraSolicitud = Format(Time, "hh:mm")
           oDoCita.IdUsuarioAuditoria = ml_idUsuario
           mo_ReglasDeProgMedica.CitasActualizaDatosXpaciente oConexion, mo_lcNombrePc, cmbIdResponsable.Text, _
                                 cmbIdServicioCE.Text, oDoCita, Val(txtNcuenta.Text), txtNombrePaciente.Text
           oConexion.CommitTrans
        End If
        
        oReglasArchivoClinico.HistoriasSolicitadasNOexistentes oConexion
        oReglasArchivoClinico.CitasNOexistentes oConexion
        oReglasArchivoClinico.ItemsBoletaFarmaciaConCantidadCERO oConexion
        If oRsTmp1!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
           ActualizaMedicoEnFuasYaEmitidas oConexionExterna, oRsTmp1!idCuentaAtencion, Val(mo_cmbIdResponsableNewPac.BoundText), _
                                        Me.txtFechaNewPac.Text, Me.txtHoraPac.Text
        End If
        If txtFechaAtencionPac.Text <> txtFechaNewPac.Text Then
           EnviaEmail oConexion, oRsTmp1!idPaciente, txtFechaNewPac.Text & " " & Me.txtHoraPac.Text, txtFechaAtencionPac.Text, oRsTmp1!idCuentaAtencion
        End If
        oConexion.Close
        oConexionExterna.Close
        
        Me.MousePointer = 1
        LimpiarDatos
        txtNcuenta.SetFocus
    End If
    Set oRsCita = Nothing
    Set oRsTmp1 = Nothing
    Set oRsTmp = Nothing
    Set oRsTmp2 = Nothing
    Set oDoCita = Nothing
    Set oCitas = Nothing
    Set oConexion = Nothing
    Set oReglasArchivoClinico = Nothing
    Set rsCitas = Nothing
    Set oRsTmp3 = Nothing
    Set oConexionExterna = Nothing
    Exit Sub
ErrorProceso1:
    oConexion.RollbackTrans
    If lnAtencionesFarmaciaServicios > 0 Then
       MsgBox "No se puede REPROGRAMAR, ya que hubo Atenciones en Farmacia/Imagenes/Laboratorio para algunos Pacientes: " & Trim(str(lnAtencionesFarmaciaServicios)), vbInformation, Me.Caption
    End If
    MsgBox Err.Description
End Sub



Private Sub Form_Load()
       mo_Formulario.HabilitarDeshabilitar cmbIdResponsablePac, False
       mo_Formulario.HabilitarDeshabilitar cmbIdServicioCEpac, False
       mo_Formulario.HabilitarDeshabilitar txtFechaAtencionPac, False
       mo_Formulario.HabilitarDeshabilitar txtNombrePaciente, False
       mo_Formulario.HabilitarDeshabilitar txtDatosDeCuenta, False
       mo_Formulario.HabilitarDeshabilitar txtPlan, False
       mo_Formulario.HabilitarDeshabilitar txtHoraFinPac, False
       
       Dim oBuscaMedicos As New SIGHNegocios.ReglasDeProgMedica
       Dim mo_AdminServHosp As New ReglasServiciosHosp
       mo_ReglasComunes.LlenaListBoxConTablaMensajesEnVentana cmbConsideraciones, "HerrReprogramacMedica"
       '
       Set mo_cmbIdResponsableNew.MiComboBox = cmbIdResponsableNew
       Set mo_cmbIdResponsableNewPac.MiComboBox = Me.cmbIdResponsableNewPac
       '
       Set mo_cmbIdResponsable.MiComboBox = cmbIdResponsable
       mo_cmbIdResponsable.BoundColumn = "IdMedico"
       mo_cmbIdResponsable.ListField = "Dmedico"
       Set mo_cmbIdResponsable.RowSource = oBuscaMedicos.MedicosSeleccionarTodosOrdenadoAlfabeticamente
       
       Me.txtFechaInicio.Text = Format(Date, sighentidades.DevuelveFechaSoloFormato_DMY)
       
       Set mo_cmbIdServicioCE.MiComboBox = cmbIdServicioCE
       mo_cmbIdServicioCE.BoundColumn = "idServicio"
       mo_cmbIdServicioCE.ListField = "descripcionLarga"
       Set mo_cmbIdServicioCE.RowSource = mo_AdminServHosp.ServiciosSeleccionarPorTipoV2(1, sghFiltraSoloActivos)
       '
       Set mo_cmbIdResponsablePac.MiComboBox = cmbIdResponsablePac
       mo_cmbIdResponsablePac.BoundColumn = "IdMedico"
       mo_cmbIdResponsablePac.ListField = "Dmedico"
       Set mo_cmbIdResponsablePac.RowSource = oBuscaMedicos.MedicosSeleccionarTodosOrdenadoAlfabeticamente
       '
       Set mo_cmbIdServicioCEpac.MiComboBox = cmbIdServicioCEpac
       mo_cmbIdServicioCEpac.BoundColumn = "idServicio"
       mo_cmbIdServicioCEpac.ListField = "descripcionLarga"
       Set mo_cmbIdServicioCEpac.RowSource = mo_AdminServHosp.ServiciosSeleccionarPorTipoV2(1, sghFiltraSoloActivos)
       '
       txtHrInicio.Text = "00:01"
       txtHrFin.Text = "23:59"
       lcParametro523 = lcBuscaParametro.SeleccionaFilaParametro(523)
       lcParametro524 = lcBuscaParametro.SeleccionaFilaParametro(524)
       lcParametro205 = lcBuscaParametro.SeleccionaFilaParametro(205)
End Sub

Private Sub txtFechaAtencionPac_LostFocus()
If Not EsFecha(txtFechaAtencionPac.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFechaAtencionPac.Text = sighentidades.FECHA_VACIA_DMY
        Exit Sub
    End If
End Sub

Private Sub txtFechaInicio_LostFocus()
If Not EsFecha(txtFechaInicio.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFechaInicio.Text = sighentidades.FECHA_VACIA_DMY
        Exit Sub
    End If
End Sub

Private Sub txtFechaNewPac_KeyDown(KeyCode As Integer, Shift As Integer)
     mo_Teclado.RealizarNavegacion KeyCode, cmbIdResponsableNewPac
End Sub

Private Sub txtFechaNewPac_LostFocus()
If Not EsFecha(txtFechaNewPac.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFechaNewPac.Text = sighentidades.FECHA_VACIA_DMY
        Exit Sub
    End If
    LlenaComboConHoraInicio
End Sub

Private Sub txtFechaRequeridaDesde_LostFocus()
If Not EsFecha(txtFechaRequeridaDesde.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFechaRequeridaDesde.Text = sighentidades.FECHA_VACIA_DMY
        Exit Sub
    End If
End Sub

Private Sub txtHoraFinPac_LostFocus()
If Not sighentidades.ValidaHora(txtHoraFinPac.Text) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
             txtHoraFinPac.Text = sighentidades.HORA_VACIA_HM
        End If
End Sub

Private Sub txtHoraPac_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtHoraPac
End Sub

Private Sub txtHoraPac_LostFocus()
'        If Not sighentidades.ValidaHora(txtHoraPac) Then
'            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
'             txtHoraPac = sighentidades.HORA_VACIA_HM
'
'        End if
        If sighentidades.ValidaHora(txtHoraPac) And txtHoraPac.Text <> "" Then
             txtHoraFinPac.Text = mo_ReglasDeProgMedica.ConvertirAHora(mo_ReglasDeProgMedica.ConvertirAMinutos(txtHoraPac.Text) + Val(lcTiempoAtencion))
        End If
End Sub

Private Sub txtHrFin_LostFocus()
 If Not sighentidades.ValidaHora(txtHrFin) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
             txtHrFin = sighentidades.HORA_VACIA_HM
        End If
End Sub

Private Sub txtHrInicio_LostFocus()
 If Not sighentidades.ValidaHora(txtHrInicio) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
             txtHrInicio = sighentidades.HORA_VACIA_HM
        End If
End Sub

Private Sub txtNcuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNcuenta
End Sub

Private Sub txtNcuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txtNcuenta_LostFocus
    End If
End Sub

Private Sub txtNcuenta_LostFocus()
   If Val(txtNcuenta.Text) > 0 Then
       Dim oRsTmp As New Recordset
       Dim lbSigue As Boolean
       Dim oConexion As New Connection
       oConexion.Open sighentidades.CadenaConexion
       oConexion.CursorLocation = adUseClient
       Set oRsTmp = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(txtNcuenta.Text, oConexion)
       txtDatosDeCuenta.Text = ""
       txtPlan.Text = ""
       txtNombrePaciente.Text = ""
       If oRsTmp.RecordCount > 0 Then
            If oRsTmp.Fields!IdEstado <> 1 Then
               MsgBox "Esa cuenta no se  encuentra ABIERTA", vbInformation, "Mensaje"
            Else
                txtDatosDeCuenta.Text = "F.Ing: " & oRsTmp.Fields!FechaIngreso & " - " & IIf(oRsTmp.Fields!idTipoServicio = 1, "Consultorios Externos", IIf(oRsTmp.Fields!idTipoServicio = 3, "Hospitalización", "Emergencia")) & " - (Est: " & Trim(oRsTmp.Fields!estadoCta) & ")"
                txtPlan.Text = "IAFA Act.: " & oRsTmp.Fields!dFuenteFinanciamiento
                txtNombrePaciente.Text = oRsTmp.Fields!NroHistoriaClinica & " " & Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & " " & oRsTmp.Fields!PrimerNombre
                txtFechaAtencionPac.Text = oRsTmp!FechaIngreso
                mo_cmbIdResponsablePac.BoundText = Trim(str(oRsTmp!IdMedicoIngreso))
                mo_cmbIdServicioCEpac.BoundText = Trim(str(oRsTmp!IdServicioIngreso))
                cmbIdServicioCEpac_Click
            End If
       End If
       oRsTmp.Close
       Set oRsTmp = Nothing
       oConexion.Close
       Set oConexion = Nothing
   End If
End Sub

Sub LimpiarDatos()
    txtNcuenta.Text = ""
    txtNombrePaciente.Text = ""
    txtDatosDeCuenta.Text = ""
    txtPlan.Text = ""
    txtHoraPac.Clear
    lcTiempoAtencion = ""
    txtHoraFinPac.Text = sighentidades.HORA_VACIA_HM
    
End Sub

Function CitasSeleccionarPorMedicoYFecha(lnIdMedico As Long, lcFecha As String) As Recordset
'        Dim lcSql As String
        Dim oRsTmp1 As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection


        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
          
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "CitasSeleccionarPorMedicoYFecha"
            Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, lnIdMedico): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, Format(lcFecha, sighentidades.DevuelveFechaSoloFormato_DMY)): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
            Set oRsTmp1.ActiveConnection = Nothing
        End With
          
        Set CitasSeleccionarPorMedicoYFecha = oRsTmp1
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
End Function




Sub LlenaComboConHoraInicio()
    If EsFecha(txtFechaNewPac.Text, "DD/MM/AAAA") = True And cmbIdResponsableNewPac.Text <> "" Then
       Dim oRsTmp1 As New Recordset
       Dim oRsTmp2 As New Recordset
       Dim oRsCitaBloqueada As New Recordset
       Dim oConexion As New Connection
       Dim lcHoraInicioCita As String, lbCitaVacia As Boolean, lnYaCitados As Long
       txtHoraPac.Clear
       
       sighentidades.AbreConexionSIGH oConexion
       Set oRsCitaBloqueada = mo_ReglasDeProgMedica.CitasBloqueadasTodas(oConexion)   '*****:nuevo PA para filtrar por FEcha,IdSERVICIO
       oRsCitaBloqueada.Filter = "fecha='" & txtFechaNewPac.Text & "' and idMedico=" & mo_cmbIdResponsableNewPac.BoundText
       
       Set oRsTmp1 = mo_ReglasDeProgMedica.ProgramacionMedicaSeleccionarPorMedicoFechaServicio(Val(mo_cmbIdResponsableNewPac.BoundText), _
                                                            txtFechaNewPac.Text, 0)
       Set oRsTmp2 = CitasSeleccionarPorMedicoYFecha(Val(mo_cmbIdResponsableNewPac.BoundText), txtFechaNewPac.Text)
       lnYaCitados = oRsTmp2.RecordCount
       If oRsTmp1.RecordCount > 0 Then
          oRsTmp1.MoveFirst
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
                    If oRsCitaBloqueada.RecordCount > 0 Then
                       oRsCitaBloqueada.MoveFirst
                       oRsCitaBloqueada.Find "HoraInicio='" & lcHoraInicioCita & "'"
                       If Not oRsCitaBloqueada.EOF Then
                          lbCitaVacia = False
                       End If
                    End If
                End If
                
                If lbCitaVacia = True Then
                   txtHoraPac.AddItem lcHoraInicioCita
                End If
                lcHoraInicioCita = mo_ReglasDeProgMedica.ConvertirAHora(mo_ReglasDeProgMedica.ConvertirAMinutos(lcHoraInicioCita) + Val(lcTiempoAtencion))
                If lcHoraInicioCita >= oRsTmp1!HoraFin Then
                   Exit Do
                End If
             Loop
             oRsTmp1.MoveNext
          Loop
       End If
       oRsTmp1.Close
       oConexion.Close
       Set oRsTmp1 = Nothing
       Set oRsTmp2 = Nothing
       Set oRsCitaBloqueada = Nothing
       Set oConexion = Nothing
    End If
End Sub

Sub ActualizaMedicoEnFuasYaEmitidas(oConexion As Connection, lnIdCuentaAtencion As Long, lnIdMedico As Long, _
                                    FechaCita As Date, HoraCita As String)
    
    Dim lcMensajeLicencia As String
'    If False Then    'licencia
'       Exit Sub
'    End If
    
    On Error GoTo ErrActFua11
    Dim oDoSisFuaAtencion As New SIGHSis.DoSisFuaAtencion
    Dim oSisFuaAtencion As New SIGHSis.SisFuaAtencion
    Dim oRsTmp771 As New Recordset
    Dim oRsTmp772 As New Recordset
    Dim lnIdEmpleado As Long
    Set oSisFuaAtencion.Conexion = oConexion
    oDoSisFuaAtencion.idCuentaAtencion = lnIdCuentaAtencion
    oDoSisFuaAtencion.IdUsuarioAuditoria = sighentidades.Usuario
    If oSisFuaAtencion.SeleccionarPorId(oDoSisFuaAtencion) = True Then
       If lnIdMedico > 0 Then
            Set oRsTmp771 = mo_ReglasDeProgMedica.MedicosSeleccionarXIdMedico(lnIdMedico)
            If oRsTmp771.RecordCount > 0 Then
                      oDoSisFuaAtencion.MedicoDocumentoTipo = oRsTmp771!idTipoDocumento
                      oDoSisFuaAtencion.FuaMedico = Left(Trim(oRsTmp771!ApellidoPaterno) & " " & Trim(oRsTmp771!ApellidoMaterno) & oRsTmp771!Nombres, 120)
                      oDoSisFuaAtencion.FuaMedicoDNI = oRsTmp771!dni
                      oDoSisFuaAtencion.FuaMedicoTipo = IIf(IsNull(oRsTmp771!TipoEmpleadoSIS), "", oRsTmp771!TipoEmpleadoSIS)
            End If
            oRsTmp771.Close
       End If
       If FechaCita <> 0 Then
          oDoSisFuaAtencion.FuaAtencionFecha = FechaCita
       End If
       If HoraCita <> "" Then
          oDoSisFuaAtencion.FuaAtencionHora = HoraCita
       End If
       oDoSisFuaAtencion.IdUsuarioAuditoria = sighentidades.Usuario
       If oSisFuaAtencion.Modificar(oDoSisFuaAtencion) = True Then
       End If
    End If
ErrActFua11:
    Set oDoSisFuaAtencion = Nothing
    Set oSisFuaAtencion = Nothing
    Set oRsTmp771 = Nothing
    Set oRsTmp772 = Nothing
    'Resume
End Sub


Sub EnviaEmail(oConexion As Connection, lnIdPaciente As Long, lcNuevaFecha As String, lcFechaActual As String, _
               lnIdCuentaAtencion As Long)
    On Error GoTo ErrEmail
    Dim mo_Pacientes As New DOPaciente
    Dim mo_email As New Procesos
    Dim lcMensajeDelAsunto As String
    Dim lcAsunto As String
    Set mo_Pacientes = mo_ReglasAdmision.PacientesSeleccionarPorId(lnIdPaciente, oConexion)
    If Not mo_Pacientes Is Nothing Then
       If mo_Pacientes.email <> "" Then
          lcAsunto = "Cita " & lcFechaActual & " a cambiado en: " & lcParametro205
          lcMensajeDelAsunto = "Se le ha programado para el: " & lcNuevaFecha & " (N° Cuenta: " & _
                               Trim(str(lnIdCuentaAtencion)) & ") "
          mo_email.EnviaEmail lcParametro524, lcParametro523, _
                              lcAsunto, _
                              "", mo_Pacientes.email, lcMensajeDelAsunto
       End If
       '
       lcMensajeDelAsunto = "Su cita del " & lcFechaActual & " a cambiado a " & lcNuevaFecha & " (N° Cuenta: " & _
                            Trim(str(lnIdCuentaAtencion)) & ") en " & lcParametro205
       mo_email.MensajeCelularEnviar mo_Pacientes, lnIdCuentaAtencion, lcMensajeDelAsunto, "REPROGRAMACION_CITA", oConexion
       '
    End If
    
ErrEmail:
    Set mo_Pacientes = Nothing
End Sub
