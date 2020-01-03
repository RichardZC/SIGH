VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Sunasa 
   Caption         =   "Genera Trama Sunasa"
   ClientHeight    =   6480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14505
   Icon            =   "Sunasa.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   14505
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar progresSunasaDetalle 
      Height          =   270
      Left            =   90
      TabIndex        =   64
      Top             =   4995
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar progresSunasa 
      Height          =   315
      Left            =   5970
      TabIndex        =   63
      Top             =   4560
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame 
      Height          =   1545
      Index           =   14
      Left            =   9690
      TabIndex        =   53
      Top             =   120
      Width           =   4695
      Begin VB.Frame Frame 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   15
         Left            =   150
         TabIndex        =   55
         Top             =   375
         Width           =   4455
         Begin VB.TextBox txtFrecurFinal 
            Height          =   315
            Left            =   3240
            TabIndex        =   57
            Text            =   "31/12/2013"
            Top             =   290
            Width           =   1035
         End
         Begin VB.TextBox txtFrecurInicial 
            Height          =   315
            Left            =   1080
            TabIndex        =   56
            Text            =   "01/12/2013"
            Top             =   290
            Width           =   1035
         End
         Begin VB.Label Label16 
            Caption         =   "Fecha Fin"
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
            Left            =   2400
            TabIndex        =   59
            Top             =   290
            Width           =   1065
         End
         Begin VB.Label Label14 
            Caption         =   "Fecha Inicio"
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
            Left            =   120
            TabIndex        =   58
            Top             =   290
            Width           =   1065
         End
      End
      Begin VB.CheckBox ckbRecursos 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Recursos de Salud"
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
         Left            =   165
         TabIndex        =   54
         Top             =   -45
         Width           =   4455
      End
   End
   Begin VB.Frame Frame 
      Height          =   1575
      Index           =   12
      Left            =   9750
      TabIndex        =   46
      Top             =   1695
      Width           =   4695
      Begin VB.CheckBox chbProcedimientos 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Procedimientos"
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
         Left            =   150
         TabIndex        =   52
         Top             =   -90
         Width           =   4425
      End
      Begin VB.Frame Frame 
         Caption         =   "Fecha de registro CPT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   13
         Left            =   135
         TabIndex        =   47
         Top             =   360
         Width           =   4455
         Begin VB.TextBox txtFcptInicio 
            Height          =   315
            Left            =   1080
            TabIndex        =   49
            Text            =   "01/12/2013"
            Top             =   290
            Width           =   1035
         End
         Begin VB.TextBox txtFcptFin 
            Height          =   315
            Left            =   3240
            TabIndex        =   48
            Text            =   "31/12/2013"
            Top             =   290
            Width           =   1035
         End
         Begin VB.Label Label13 
            Caption         =   "Fecha Inicio"
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
            Left            =   120
            TabIndex        =   51
            Top             =   290
            Width           =   1065
         End
         Begin VB.Label Label12 
            Caption         =   "Fecha Fin"
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
            Left            =   2400
            TabIndex        =   50
            Top             =   290
            Width           =   1065
         End
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "*No se tomará en cuenta los SERVICIOS sin código SUSALUD "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   120
         TabIndex        =   66
         Top             =   1200
         Width           =   4500
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   135
      TabIndex        =   42
      Top             =   5310
      Width           =   14310
      Begin VB.CommandButton cmdExportar2016 
         Caption         =   "Exporta Trama desde 2016"
         DisabledPicture =   "Sunasa.frx":000C
         DownPicture     =   "Sunasa.frx":046C
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5582
         Picture         =   "Sunasa.frx":08E1
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   210
         Width           =   1665
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "Sunasa.frx":0D56
         DownPicture     =   "Sunasa.frx":121A
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7377
         Picture         =   "Sunasa.frx":1706
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   210
         Width           =   1575
      End
      Begin VB.CommandButton cmdExportaTrama 
         Caption         =   "Exporta Trama hasta 2015"
         DisabledPicture =   "Sunasa.frx":1BF2
         DownPicture     =   "Sunasa.frx":2052
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         Picture         =   "Sunasa.frx":24C7
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   225
         Visible         =   0   'False
         Width           =   1665
      End
   End
   Begin VB.Frame Frame 
      Height          =   1620
      Index           =   10
      Left            =   4935
      TabIndex        =   35
      Top             =   1725
      Width           =   4695
      Begin VB.Frame Frame 
         Caption         =   "Fecha de egreso Hospitalización"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   11
         Left            =   150
         TabIndex        =   37
         Top             =   375
         Width           =   4455
         Begin VB.TextBox txtFechaFinPresSaludHospitalizacion 
            Height          =   315
            Left            =   3240
            TabIndex        =   39
            Text            =   "31/12/2013"
            Top             =   290
            Width           =   1035
         End
         Begin VB.TextBox txtFechaInicioPresSaludHospitalizacion 
            Height          =   315
            Left            =   1080
            TabIndex        =   38
            Text            =   "01/12/2013"
            Top             =   290
            Width           =   1035
         End
         Begin VB.Label Label11 
            Caption         =   "Fecha Fin"
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
            Left            =   2400
            TabIndex        =   41
            Top             =   290
            Width           =   1065
         End
         Begin VB.Label Label10 
            Caption         =   "Fecha Inicio"
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
            Left            =   120
            TabIndex        =   40
            Top             =   290
            Width           =   1065
         End
      End
      Begin VB.CheckBox chbPresSaludhospitalizacion 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Hospitalización (Prod.Asistencial y Morb)"
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
         Left            =   120
         TabIndex        =   36
         Top             =   -75
         Width           =   4455
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "*No se tomará en cuenta los SERVICIOS sin código SUSALUD "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   105
         TabIndex        =   65
         Top             =   1155
         Width           =   4500
      End
   End
   Begin VB.Frame Frame 
      Height          =   1140
      Index           =   8
      Left            =   4920
      TabIndex        =   28
      Top             =   3360
      Width           =   4695
      Begin VB.CheckBox chbPresSaludHospParto 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Parto"
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
         Left            =   330
         TabIndex        =   34
         Top             =   -75
         Width           =   4185
      End
      Begin VB.Frame Frame 
         Caption         =   "Fecha de Ingreso Emerg/Hosp"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   9
         Left            =   90
         TabIndex        =   29
         Top             =   330
         Width           =   4455
         Begin VB.TextBox txtFechaInicioPresSaludHospParto 
            Height          =   315
            Left            =   1080
            TabIndex        =   31
            Text            =   "01/12/2013"
            Top             =   290
            Width           =   1035
         End
         Begin VB.TextBox txtFechaFinPresSaludHospParto 
            Height          =   315
            Left            =   3240
            TabIndex        =   30
            Text            =   "31/12/2013"
            Top             =   290
            Width           =   1035
         End
         Begin VB.Label Label9 
            Caption         =   "Fecha Inicio"
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
            Left            =   120
            TabIndex        =   33
            Top             =   290
            Width           =   1065
         End
         Begin VB.Label Label8 
            Caption         =   "Fecha Fin"
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
            Left            =   2400
            TabIndex        =   32
            Top             =   290
            Width           =   1065
         End
      End
   End
   Begin VB.Frame Frame 
      Height          =   1185
      Index           =   6
      Left            =   120
      TabIndex        =   21
      Top             =   3360
      Width           =   4695
      Begin VB.Frame Frame 
         Caption         =   "Fecha de Ingreso CE/Emerg"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   7
         Left            =   135
         TabIndex        =   23
         Top             =   360
         Width           =   4455
         Begin VB.TextBox txtFreferFin 
            Height          =   315
            Left            =   3240
            TabIndex        =   25
            Text            =   "31/12/2013"
            Top             =   290
            Width           =   1035
         End
         Begin VB.TextBox txtFreferInicio 
            Height          =   315
            Left            =   1080
            TabIndex        =   24
            Text            =   "01/12/2013"
            Top             =   290
            Width           =   1035
         End
         Begin VB.Label Label7 
            Caption         =   "Fecha Fin"
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
            Left            =   2400
            TabIndex        =   27
            Top             =   290
            Width           =   1065
         End
         Begin VB.Label Label6 
            Caption         =   "Fecha Inicio"
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
            Left            =   120
            TabIndex        =   26
            Top             =   290
            Width           =   1065
         End
      End
      Begin VB.CheckBox chbReferencias 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Referencias"
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
         Left            =   150
         TabIndex        =   22
         Top             =   -90
         Width           =   4425
      End
   End
   Begin VB.Frame Frame 
      Height          =   1365
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   1980
      Width           =   4695
      Begin VB.CheckBox chbPresSaludEmergencia 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Emergencia (Prod.Asistencial y Morb)"
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
         Left            =   195
         TabIndex        =   20
         Top             =   -90
         Width           =   4380
      End
      Begin VB.Frame Frame 
         Caption         =   "Fecha de egreso a Emergencia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   5
         Left            =   75
         TabIndex        =   15
         Top             =   390
         Width           =   4455
         Begin VB.TextBox txtFechaInicioPresSaludEmergencia 
            Height          =   315
            Left            =   1080
            TabIndex        =   17
            Text            =   "01/12/2013"
            Top             =   290
            Width           =   1035
         End
         Begin VB.TextBox txtFechaFinPresSaludEmergencia 
            Height          =   315
            Left            =   3240
            TabIndex        =   16
            Text            =   "31/12/2013"
            Top             =   290
            Width           =   1035
         End
         Begin VB.Label Label5 
            Caption         =   "Fecha Inicio"
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
            Left            =   120
            TabIndex        =   19
            Top             =   290
            Width           =   1065
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Fin"
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
            Left            =   2400
            TabIndex        =   18
            Top             =   290
            Width           =   1065
         End
      End
   End
   Begin VB.Frame Frame 
      Height          =   1560
      Index           =   2
      Left            =   4920
      TabIndex        =   7
      Top             =   120
      Width           =   4695
      Begin VB.Frame Frame 
         Caption         =   "Fecha de Citas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   105
         TabIndex        =   9
         Top             =   420
         Width           =   4455
         Begin VB.TextBox txtFechaFinEmisionCitas 
            Height          =   315
            Left            =   3240
            TabIndex        =   11
            Text            =   "31/12/2013"
            Top             =   290
            Width           =   1035
         End
         Begin VB.TextBox txtfechaInicialEmisionCitas 
            Height          =   315
            Left            =   1080
            TabIndex        =   10
            Text            =   "01/12/2013"
            Top             =   290
            Width           =   1035
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Fin"
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
            Left            =   2400
            TabIndex        =   13
            Top             =   290
            Width           =   1065
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Inicio"
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
            Left            =   120
            TabIndex        =   12
            Top             =   290
            Width           =   1065
         End
      End
      Begin VB.CheckBox chbEmisionCitas 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Consultorios Externos (Prod.Asistencial y Morb)"
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
         TabIndex        =   8
         Top             =   -75
         Width           =   4305
      End
   End
   Begin VB.Frame Frame 
      Height          =   1770
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.CheckBox chbProgAsistencial 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Programación Asistencial"
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
         Left            =   135
         TabIndex        =   6
         Top             =   -75
         Width           =   4455
      End
      Begin VB.Frame Frame 
         Caption         =   "Fecha Programacion Médica"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   375
         Width           =   4455
         Begin VB.TextBox txtFProgInicio 
            Height          =   315
            Left            =   1140
            TabIndex        =   3
            Text            =   "01/12/2013"
            Top             =   290
            Width           =   1035
         End
         Begin VB.TextBox txtFprogFinal 
            Height          =   315
            Left            =   3240
            TabIndex        =   2
            Text            =   "31/12/2013"
            Top             =   290
            Width           =   1035
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Inicio"
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
            TabIndex        =   5
            Top             =   300
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Fin"
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
            Left            =   2415
            TabIndex        =   4
            Top             =   315
            Width           =   765
         End
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "* Algunos TURNOS deben tener TIPO ACTIVIDAD"
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
         Left            =   150
         TabIndex        =   62
         Top             =   1395
         Width           =   4155
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "* Chequear parametros 503 y 504"
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
         Left            =   165
         TabIndex        =   61
         Top             =   1140
         Width           =   2820
      End
   End
   Begin VB.Label lblTabla 
      Caption         =   "...."
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
      TabIndex        =   60
      Top             =   4635
      Width           =   5370
   End
End
Attribute VB_Name = "Sunasa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Genera datos para SUNASA
'        Programado por: Palomino Y
'        Fecha: Enero 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Dim oRsPatologia As New Recordset
Dim oRsFarmacia As New Recordset
Const lnIdUsuario As Long = 738
Dim ml_idUsuario As Long
Dim mo_lcNombrePc  As String
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
Dim mo_ReglasHoteleria As New SIGHNegocios.ReglasHoteleria
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property

Property Let idUsuario(lIdValue As Long)
    ml_idUsuario = lIdValue
End Property

Private Sub btnCancelar_Click()
Me.Visible = False
End Sub





Private Sub cmdExportar2016_Click()

If wxFranklin = "*" Then Exit Sub

Dim lcMensajeInformacion As String
Dim lbSelecciono As Boolean
Dim oConexionFox As New ADODB.Connection
Dim oConexion As New ADODB.Connection
Dim objShell
Dim lnMaximoTramas As Integer
Dim lnContadoTramas As Integer
Dim lcRutaExportar As String
Dim lcIpress As String, lcUgipress As String

lcIpress = Right("00000000" & lcBuscaParametro.SeleccionaFilaParametro(208), 8)
lcUgipress = lcIpress
lcRutaExportar = lcBuscaParametro.SeleccionaFilaParametro(313)

oConexion.CursorLocation = adUseClient
oConexion.CommandTimeout = 300
oConexion.Open sighentidades.CadenaConexion

oConexionFox.CommandTimeout = 300
oConexionFox.Open "DSN=his"



lnMaximoTramas = 0
If Me.chbPresSaludHospParto.Value = 1 Then lnMaximoTramas = lnMaximoTramas + 1
If Me.chbEmisionCitas.Value = 1 Then lnMaximoTramas = lnMaximoTramas + 1
If Me.chbPresSaludEmergencia.Value = 1 Then lnMaximoTramas = lnMaximoTramas + 1
If Me.chbPresSaludhospitalizacion.Value = 1 Then lnMaximoTramas = lnMaximoTramas + 1
If chbProcedimientos.Value = 1 Then lnMaximoTramas = lnMaximoTramas + 1
If chbReferencias.Value = 1 Then lnMaximoTramas = lnMaximoTramas + 1
If chbProgAsistencial.Value = 1 Then lnMaximoTramas = lnMaximoTramas + 1
If lnMaximoTramas > 0 Then
    progresSunasa.Min = 0
    progresSunasa.Max = lnMaximoTramas
End If
lnContadoTramas = 0
progresSunasa.Value = lnContadoTramas

lcMensajeInformacion = "Se generó las tramas:"
lbSelecciono = False

If Me.chbEmisionCitas.Value = 1 Then
    lbSelecciono = True
    lcMensajeInformacion = lcMensajeInformacion & " " & chbEmisionCitas.Caption & ", "
    GeneraTramaEmisionCitas2016 oConexion, lcIpress, lcUgipress, lcRutaExportar
    lnContadoTramas = lnContadoTramas + 1
    progresSunasa.Value = lnContadoTramas
End If
If Me.chbPresSaludEmergencia.Value = 1 Then
    lbSelecciono = True
    lcMensajeInformacion = lcMensajeInformacion & " " & chbPresSaludEmergencia.Caption & ", "
    GeneraTramaEmisionEmergencia2016 oConexion, lcIpress, lcUgipress, lcRutaExportar
    lnContadoTramas = lnContadoTramas + 1
    progresSunasa.Value = lnContadoTramas
End If
If Me.chbPresSaludhospitalizacion.Value = 1 Then
    lbSelecciono = True
    lcMensajeInformacion = lcMensajeInformacion & " " & chbPresSaludhospitalizacion.Caption & ", "
    GeneraTramaEmisionHospitalizacion2016 oConexion, lcIpress, lcUgipress, lcRutaExportar
    lnContadoTramas = lnContadoTramas + 1
    progresSunasa.Value = lnContadoTramas
End If
If Me.chbPresSaludHospParto.Value = 1 Then
    lbSelecciono = True
    lcMensajeInformacion = lcMensajeInformacion & " " & chbPresSaludHospParto.Caption & ", "
    GeneraTramaPrestacionesSaludHospParto2016 oConexion, lcIpress, lcUgipress, lcRutaExportar
    lnContadoTramas = lnContadoTramas + 1
    progresSunasa.Value = lnContadoTramas
End If
If chbProcedimientos.Value = 1 Then
    lbSelecciono = True
    lcMensajeInformacion = lcMensajeInformacion & " " & chbProcedimientos.Caption & ", "
    GeneraTramaPrestacionesSaludCPT2016 oConexion, lcIpress, lcUgipress, lcRutaExportar
    lnContadoTramas = lnContadoTramas + 1
    progresSunasa.Value = lnContadoTramas
End If
If chbReferencias.Value = 1 Then
    lbSelecciono = True
    lcMensajeInformacion = lcMensajeInformacion & " " & chbReferencias.Caption & ", "
    GeneraTramaReferencias2016 oConexion, lcIpress, lcUgipress, lcRutaExportar
    lnContadoTramas = lnContadoTramas + 1
    progresSunasa.Value = lnContadoTramas
End If
If chbProgAsistencial.Value = 1 Then
    lbSelecciono = True
    lcMensajeInformacion = lcMensajeInformacion & " " & chbProgAsistencial.Caption & ", "
    GeneraTramaProgramacion2016 oConexion, lcIpress, lcUgipress, lcRutaExportar
    lnContadoTramas = lnContadoTramas + 1
    progresSunasa.Value = lnContadoTramas
End If

Me.Refresh
oConexion.Close
oConexionFox.Close
Set oConexion = Nothing
Set oConexionFox = Nothing
If lbSelecciono = True Then
   
    MsgBox lcMensajeInformacion, vbInformation, "SUNASA_TRAMAS"
    Set objShell = CreateObject("Shell.Application")
    objShell.Open lcRutaExportar
    MsgBox "Se exportó archivos txt SUNASA en: " & lcRutaExportar
Else
    MsgBox "Elija una trama a generar", vbInformation, "SUNASA_TRAMAS"
End If

End Sub

Private Sub cmdExportaTrama_Click()
Dim lcMensajeInformacion As String
Dim lbSelecciono As Boolean
Dim oConexionFox As New ADODB.Connection
Dim oConexion As New ADODB.Connection
Dim objShell
Dim lnMaximoTramas As Integer
Dim lnContadoTramas As Integer
Dim lcRutaExportar As String
lcRutaExportar = lcBuscaParametro.SeleccionaFilaParametro(313)

oConexion.CursorLocation = adUseClient
oConexion.CommandTimeout = 300
oConexion.Open sighentidades.CadenaConexion

oConexionFox.CommandTimeout = 300
oConexionFox.Open "DSN=his"



lnMaximoTramas = 0
If Me.chbProgAsistencial.Value = 1 Then lnMaximoTramas = lnMaximoTramas + 1
If Me.chbEmisionCitas.Value = 1 Then lnMaximoTramas = lnMaximoTramas + 1
If Me.chbReferencias.Value = 1 Then lnMaximoTramas = lnMaximoTramas + 1
If Me.chbPresSaludEmergencia.Value = 1 Then lnMaximoTramas = lnMaximoTramas + 1
If Me.chbPresSaludhospitalizacion.Value = 1 Then lnMaximoTramas = lnMaximoTramas + 1
If Me.chbPresSaludHospParto.Value = 1 Then lnMaximoTramas = lnMaximoTramas + 1

If lnMaximoTramas > 0 Then
    progresSunasa.Min = 0
    progresSunasa.Max = lnMaximoTramas
End If
lnContadoTramas = 0
progresSunasa.Value = lnContadoTramas

lcMensajeInformacion = "Se generó las tramas:"
lbSelecciono = False
If Me.chbProgAsistencial.Value = 1 Then
    lbSelecciono = True
    lcMensajeInformacion = lcMensajeInformacion & " Sunasa_ProgAsistencial,"
    GeneraTramaProgramacionAsistencial oConexion, oConexionFox
    lnContadoTramas = lnContadoTramas + 1
    progresSunasa.Value = lnContadoTramas
'    Me.Refresh
End If
If Me.chbEmisionCitas.Value = 1 Then
    lbSelecciono = True
    lcMensajeInformacion = lcMensajeInformacion & " Sunasa_EmisionCitas,"
    GeneraTramaEmisionCitas oConexion, oConexionFox
    lnContadoTramas = lnContadoTramas + 1
    progresSunasa.Value = lnContadoTramas
'    Me.Refresh
End If
If Me.chbReferencias.Value = 1 Then
    lbSelecciono = True
    lcMensajeInformacion = lcMensajeInformacion & " Sunasa_PrestacionSaludConsultorios,"
    GeneraTramaPrestacionesSaludConsultorios oConexion, oConexionFox
    lnContadoTramas = lnContadoTramas + 1
    progresSunasa.Value = lnContadoTramas
'    Me.Refresh
End If
If Me.chbPresSaludEmergencia.Value = 1 Then
    lbSelecciono = True
    lcMensajeInformacion = lcMensajeInformacion & " Sunasa_PrestacionSaludEmergencia,"
    GeneraTramaPrestacionesSaludEmergencia oConexion, oConexionFox
    lnContadoTramas = lnContadoTramas + 1
    progresSunasa.Value = lnContadoTramas
'    Me.Refresh
End If
If Me.chbPresSaludhospitalizacion.Value = 1 Then
    lbSelecciono = True
    lcMensajeInformacion = lcMensajeInformacion & " Sunasa_PrestacionSaludHospitalizacion,"
    GeneraTramaPrestacionesSaludHospitalizacion oConexion, oConexionFox
    lnContadoTramas = lnContadoTramas + 1
    progresSunasa.Value = lnContadoTramas
'    Me.Refresh
End If
If Me.chbPresSaludHospParto.Value = 1 Then
    lbSelecciono = True
    lcMensajeInformacion = lcMensajeInformacion & " Sunasa_PrestacionSaludHospParto,"
    GeneraTramaPrestacionesSaludHospParto oConexion, oConexionFox
    lnContadoTramas = lnContadoTramas + 1
    progresSunasa.Value = lnContadoTramas
'    Me.Refresh
End If
oConexion.Close
oConexionFox.Close
Set oConexion = Nothing
Set oConexionFox = Nothing
If lbSelecciono = True Then
   
    MsgBox lcMensajeInformacion, vbInformation, "SUNASA_TRAMAS"
    Set objShell = CreateObject("Shell.Application")
    objShell.Open lcRutaExportar
    MsgBox "Se exportó archivos txt SUNASA en: " & lcRutaExportar
Else
    MsgBox "Elija una trama a generar", vbInformation, "SUNASA_TRAMAS"
End If
End Sub

Sub GeneraTramaProgramacionAsistencial(oConexion As ADODB.Connection, oConexionFox As ADODB.Connection)
Dim oRsTmp As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRsFox As New Recordset
Dim lcSql As String
Dim lcLineaTxtPlano As String
Dim lnContadorDetalle As Long

'Leer datos del SISGalenPlus - PROGRAMACION ASISTENCIAL
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "Sunasa_TramaProgramacionAsistencial"
        Set oParameter = .CreateParameter("@FechaIni", adDBTimeStamp, adParamInput, 0, Format(Me.txtFProgInicio.Text, "dd/mm/yyyy")): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, Format(Me.txtFprogFinal.Text, "dd/mm/yyyy")): .Parameters.Append oParameter
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
   End With
   Set oCommand = Nothing
   Set oParameter = Nothing
   
   If oRsTmp.RecordCount > 0 Then
        progresSunasaDetalle.Min = 0
        progresSunasaDetalle.Max = oRsTmp.RecordCount
   Else
        progresSunasaDetalle.Min = 0
        progresSunasaDetalle.Max = 1
        progresSunasaDetalle.Value = 1
        Me.Refresh
   End If
   lnContadorDetalle = 0
        
'Cargar en la tabla temporal ProgAsistencial
    Dim fso
    Dim act
    Set fso = CreateObject("scripting.filesystemobject")
    lcLineaTxtPlano = ""
    Set act = fso.CreateTextFile(lcBuscaParametro.SeleccionaFilaParametro(313) & "TramaProgramacionAsistencial.txt", True)
    
    If oRsTmp.RecordCount > 0 Then
       'Inicializa tabla fox
       lcSql = "delete from su_asist.dbf" 'where Cod_ipre='" & Right("00000000000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(280)), 9) & "'"
       oRsFox.Open "select * from su_asist", oConexionFox, adOpenKeyset, adLockOptimistic
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          oRsFox.AddNew
          oRsFox.Fields!Cod_ipre = Right("00000000000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(280)), 9)
          oRsFox.Fields!PerInfPr = oRsTmp.Fields!PerInfProgAsistencial
          oRsFox.Fields!ColprofR = oRsTmp.Fields!CodigoColegio
          oRsFox.Fields!Numcoleg = oRsTmp.Fields!NumeroColegiatura
          oRsFox.Fields!ServEspB = oRsTmp.Fields!especialidad
          oRsFox.Fields!Activida = ""
          oRsFox.Fields!SubActiv = ""
          oRsFox.Fields!TipoOfer = ""
          oRsFox.Fields!FecRegAc = ""
          oRsFox.Fields!FecAperA = ""
          oRsFox.Fields!Fecnodif = ""
          oRsFox.Fields!FecSusAc = ""
          oRsFox.Fields!FecRePro = ""
          oRsFox.Fields!FecIniAs = ""
          oRsFox.Fields!HorIniTu = ""
          oRsFox.Fields!FecFinAs = ""
          oRsFox.Fields!HorFinTu = ""
          oRsFox.Update
          
            lcLineaTxtPlano = ""
            lcLineaTxtPlano = lcLineaTxtPlano & Right("00000000000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(280)), 9) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & oRsTmp.Fields!PerInfProgAsistencial & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & oRsTmp.Fields!CodigoColegio & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & oRsTmp.Fields!NumeroColegiatura & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & oRsTmp.Fields!especialidad & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "|"
            act.WriteLine (lcLineaTxtPlano)
            
            lnContadorDetalle = lnContadorDetalle + 1
            progresSunasaDetalle.Value = lnContadorDetalle
            Me.Refresh
             
          oRsTmp.MoveNext
       Loop
    End If
    If oRsFox.State = 1 Then oRsFox.Close
    act.Close
    oRsTmp.Close
End Sub

Sub GeneraTramaEmisionCitas(oConexion As ADODB.Connection, oConexionFox As ADODB.Connection)
Dim oRsTmp As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRsFox As New Recordset
Dim lcSql As String
Dim lcLineaTxtPlano As String
Dim lnContadorDetalle As Long
  
'Leer datos del SISGalenPlus - EMISION DE CITAS
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "Sunasa_TramaEmisionCitas"
        Set oParameter = .CreateParameter("@FechaIni", adDBTimeStamp, adParamInput, 0, Format(Me.txtfechaInicialEmisionCitas.Text, "dd/mm/yyyy")): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, Format(Me.txtFechaFinEmisionCitas.Text, "dd/mm/yyyy")): .Parameters.Append oParameter
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
   End With
   Set oCommand = Nothing
   Set oParameter = Nothing
   
      If oRsTmp.RecordCount > 0 Then
        progresSunasaDetalle.Min = 0
        progresSunasaDetalle.Max = oRsTmp.RecordCount
   Else
        progresSunasaDetalle.Min = 0
        progresSunasaDetalle.Max = 1
        progresSunasaDetalle.Value = 1
        Me.Refresh
   End If
   lnContadorDetalle = 0

'Cargar en la tabla temporal ProgAsistencial
    Dim fso
    Dim act
    Set fso = CreateObject("scripting.filesystemobject")
    lcLineaTxtPlano = ""
    Set act = fso.CreateTextFile(lcBuscaParametro.SeleccionaFilaParametro(313) & "TramaEmisionCitas.txt", True)
    
    If oRsTmp.RecordCount > 0 Then
       'Inicializa tabla fox
       lcSql = "delete from Su_Citas.DBF" ' where Cod_ipre='" & Right("00000000000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(280)), 9) & "'"
       oRsFox.Open "select * from Su_Citas.DBF", oConexionFox, adOpenKeyset, adLockOptimistic
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          oRsFox.AddNew
          oRsFox.Fields!Cod_ipre = Right("00000000000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(280)), 9)
          oRsFox.Fields!PerInfPr = oRsTmp.Fields!PerInfProgAsistencial
          oRsFox.Fields!ColprofR = oRsTmp.Fields!CodigoColegio
          oRsFox.Fields!Numcoleg = oRsTmp.Fields!NumeroColegiatura
          oRsFox.Fields!ServEspB = oRsTmp.Fields!especialidad
          oRsFox.Fields!Activida = ""
          oRsFox.Fields!SubActiv = ""
          oRsFox.Fields!TipoOfer = ""
          oRsFox.Update
          
            lcLineaTxtPlano = ""
            lcLineaTxtPlano = lcLineaTxtPlano & Right("00000000000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(280)), 9) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & oRsTmp.Fields!PerInfProgAsistencial & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & oRsTmp.Fields!CodigoColegio & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & oRsTmp.Fields!NumeroColegiatura & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & oRsTmp.Fields!especialidad & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "|"
            act.WriteLine (lcLineaTxtPlano)
            
            lnContadorDetalle = lnContadorDetalle + 1
            progresSunasaDetalle.Value = lnContadorDetalle
            Me.Refresh
             
          oRsTmp.MoveNext
       Loop
    End If
    If oRsFox.State = 1 Then oRsFox.Close
    act.Close
    oRsTmp.Close
End Sub

Sub GeneraTramaPrestacionesSaludConsultorios(oConexion As ADODB.Connection, oConexionFox As ADODB.Connection)
Dim oRsTmp As New Recordset
Dim oRsTmp2 As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRsFox As New Recordset
Dim lcSql As String
Dim lnNumDx As Integer
Dim lcLineaTxtPlano As String
Dim lcPrimerDx As String
Dim lcTipoPriDx As String
Dim lcSegundoDx As String
Dim lcTipoSegDx As String
Dim lcTercerDx As String
Dim lcTipoTerDx As String
Dim lnContadorDetalle As Long

'Leer datos del SISGalenPlus - PRESTACIONESSALUDCONSULTORIOS
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "Sunasa_TramaPresSaludConsultorios"
        Set oParameter = .CreateParameter("@FechaIni", adDBTimeStamp, adParamInput, 0, Format(Me.txtFreferInicio.Text, "dd/mm/yyyy")): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, Format(Me.txtFreferFin.Text, "dd/mm/yyyy")): .Parameters.Append oParameter
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
   End With
   Set oCommand = Nothing
   Set oParameter = Nothing

   If oRsTmp.RecordCount > 0 Then
        progresSunasaDetalle.Min = 0
        progresSunasaDetalle.Max = oRsTmp.RecordCount
   Else
        progresSunasaDetalle.Min = 0
        progresSunasaDetalle.Max = 1
        progresSunasaDetalle.Value = 1
        Me.Refresh
   End If
   lnContadorDetalle = 0


'Cargar en la tabla temporal Sunasa_PrestacionsSalud_Emergencia
    Dim fso
    Dim act
    Set fso = CreateObject("scripting.filesystemobject")
    lcLineaTxtPlano = ""
    Set act = fso.CreateTextFile(lcBuscaParametro.SeleccionaFilaParametro(313) & "TramaPrestacionesSaludConsultorios.txt", True)
    
    If oRsTmp.RecordCount > 0 Then
        'Inicializa tabla fox
        lcSql = "delete from su_consl.dbf" ' where Cod_ipre='" & Right("00000000000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(280)), 9) & "'"
        oRsFox.Open "select * from su_consl.dbf", oConexionFox, adOpenKeyset, adLockOptimistic
        oRsTmp.MoveFirst
        Do While Not oRsTmp.EOF
           oRsFox.AddNew
           oRsFox.Fields!Cod_ipre = Right("00000000000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(280)), 9)
           oRsFox.Fields!PeriodoR = ""
           oRsFox.Fields!CodIafas = ""
           oRsFox.Fields!NumHistC = IIf(IsNull(oRsTmp.Fields!NroHistoriaClinica), "", oRsTmp.Fields!NroHistoriaClinica)
           oRsFox.Fields!TipDocId = IIf(IsNull(oRsTmp.Fields!TipDocIdentidad), "", oRsTmp.Fields!TipDocIdentidad)
           oRsFox.Fields!NumDocId = IIf(IsNull(oRsTmp.Fields!nroDocumento), "", oRsTmp.Fields!nroDocumento)
           oRsFox.Fields!RegSegur = ""
           oRsFox.Fields!SexoPaci = IIf(IsNull(oRsTmp.Fields!idTipoSexo), "", oRsTmp.Fields!idTipoSexo)
           oRsFox.Fields!FecNacPa = IIf(IsNull(oRsTmp.Fields!FechaNacimiento), "", oRsTmp.Fields!FechaNacimiento)
           oRsFox.Fields!FecAtenc = IIf(IsNull(oRsTmp.Fields!FechaIngreso), "", oRsTmp.Fields!FechaIngreso)
           oRsFox.Fields!HoraAten = IIf(IsNull(oRsTmp.Fields!HoraIngreso), "", oRsTmp.Fields!HoraIngreso)
           oRsFox.Fields!Activida = ""
           oRsFox.Fields!ColprofR = IIf(IsNull(oRsTmp.Fields!cod_col), "", oRsTmp.Fields!cod_col)
           oRsFox.Fields!Numcoleg = IIf(IsNull(oRsTmp.Fields!Colegiatura), "", oRsTmp.Fields!Colegiatura)
           oRsFox.Fields!ServEspB = IIf(IsNull(oRsTmp.Fields!nombre), "", oRsTmp.Fields!nombre)
           oRsFox.Fields!FecEmisi = ""
           
           With oCommand
              .CommandType = adCmdStoredProc
              Set .ActiveConnection = oConexion
              .CommandTimeout = 150
              .CommandText = "Suna_TramaPresSaludConsultoriosDx"
              Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, oRsTmp.Fields!idAtencion): .Parameters.Append oParameter
              Set oRsTmp2 = .Execute
              Set oRsTmp2.ActiveConnection = Nothing
           End With
           Set oCommand = Nothing
           Set oParameter = Nothing
           'Inicializar Dx
            lcPrimerDx = ""
            lcTipoPriDx = ""
            lcSegundoDx = ""
            lcTipoSegDx = ""
            lcTercerDx = ""
            lcTipoTerDx = ""
           lnNumDx = 1
           If oRsTmp2.RecordCount > 0 Then
             oRsTmp2.MoveFirst
             Do While Not oRsTmp2.EOF
               Select Case lnNumDx
               Case 1
                   lcPrimerDx = oRsTmp2.Fields!CodigoCie2004
                   lcTipoPriDx = DevuelveCodigoTipoDx(oRsTmp2.Fields!codigo)
                   lnNumDx = lnNumDx + 1
               Case 2
                   lcSegundoDx = oRsTmp2.Fields!CodigoCie2004
                   lcTipoSegDx = DevuelveCodigoTipoDx(oRsTmp2.Fields!codigo)
                   lnNumDx = lnNumDx + 1
               Case 3
                   lcTercerDx = oRsTmp2.Fields!CodigoCie2004
                   lcTipoTerDx = DevuelveCodigoTipoDx(oRsTmp2.Fields!codigo)
                   lnNumDx = lnNumDx + 1
               End Select
               oRsTmp2.MoveNext
             Loop
           End If
           oRsFox.Fields!Primerdi = lcPrimerDx
           oRsFox.Fields!PrimerTi = lcTipoPriDx
           oRsFox.Fields!SegundoD = lcSegundoDx
           oRsFox.Fields!SegundoT = lcTipoSegDx
           oRsFox.Fields!TercerDi = lcTercerDx
           oRsFox.Fields!TercerTi = lcTipoTerDx
           oRsFox.Fields!Resultad = ""
           oRsFox.Update
          
            lcLineaTxtPlano = ""
            lcLineaTxtPlano = lcLineaTxtPlano & Right("00000000000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(280)), 9) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "" & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "" & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!NroHistoriaClinica), "", oRsTmp.Fields!NroHistoriaClinica) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!TipDocIdentidad), "", oRsTmp.Fields!TipDocIdentidad) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!nroDocumento), "", oRsTmp.Fields!nroDocumento) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "" & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!idTipoSexo), "", oRsTmp.Fields!idTipoSexo) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!FechaNacimiento), "", oRsTmp.Fields!FechaNacimiento) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!FechaIngreso), "", oRsTmp.Fields!FechaIngreso) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!FechaIngreso), "", oRsTmp.Fields!FechaIngreso) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "" & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!cod_col), "", oRsTmp.Fields!cod_col) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!Colegiatura), "", oRsTmp.Fields!Colegiatura) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!nombre), "", oRsTmp.Fields!nombre) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "" & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & lcPrimerDx & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & lcTipoPriDx & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & lcSegundoDx & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & lcTipoSegDx & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & lcTercerDx & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & lcTipoTerDx & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "" & "|"
            act.WriteLine (lcLineaTxtPlano)
            
            lnContadorDetalle = lnContadorDetalle + 1
            progresSunasaDetalle.Value = lnContadorDetalle
            Me.Refresh
            
          oRsTmp.MoveNext
       Loop
    End If
    If oRsFox.State = 1 Then oRsFox.Close
    act.Close
    oRsTmp.Close
End Sub

Sub GeneraTramaPrestacionesSaludEmergencia(oConexion As ADODB.Connection, oConexionFox As ADODB.Connection)
Dim oRsTmp As New Recordset
Dim oRsTmp2 As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRsFox As New Recordset
Dim lcSql As String
Dim lcLineaTxtPlano As String
Dim lcDx As String
Dim lcTipoDx As String
Dim lnContadorDetalle As Long

'Leer datos del SISGalenPlus - PRESTACIONESSALUDEMERGENCIA
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "Sunasa_TramaPresSaludEmergencia"
        Set oParameter = .CreateParameter("@FechaIni", adDBTimeStamp, adParamInput, 0, Format(Me.txtFechaInicioPresSaludEmergencia.Text, "dd/mm/yyyy")): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, Format(Me.txtFreferFin.Text, "dd/mm/yyyy")): .Parameters.Append oParameter
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
   End With
   Set oCommand = Nothing
   Set oParameter = Nothing

   If oRsTmp.RecordCount > 0 Then
        progresSunasaDetalle.Min = 0
        progresSunasaDetalle.Max = oRsTmp.RecordCount
   Else
        progresSunasaDetalle.Min = 0
        progresSunasaDetalle.Max = 1
        progresSunasaDetalle.Value = 1
        Me.Refresh
   End If
   lnContadorDetalle = 0


'Cargar en la tabla temporal Sunasa_PrestacionsSalud_Emergencia
    Dim fso
    Dim act
    Set fso = CreateObject("scripting.filesystemobject")
    lcLineaTxtPlano = ""
    Set act = fso.CreateTextFile(lcBuscaParametro.SeleccionaFilaParametro(313) & "TramaPrestacionesSaludEmergencia.txt", True)
    
    If oRsTmp.RecordCount > 0 Then
       'Inicializa tabla fox
       lcSql = "delete from su_emerg.dbf" ' where Cod_ipre='" & Right("00000000000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(280)), 9) & "'"
       oRsFox.Open "select * from su_emerg.dbf", oConexionFox, adOpenKeyset, adLockOptimistic
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          oRsFox.AddNew
          oRsFox.Fields!Cod_ipre = Right("00000000000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(280)), 9)
          oRsFox.Fields!PeriodoR = ""
          oRsFox.Fields!CodIafas = ""
          oRsFox.Fields!NumHistC = IIf(IsNull(oRsTmp.Fields!NroHistoriaClinica), "", oRsTmp.Fields!NroHistoriaClinica)
          oRsFox.Fields!TipDocId = IIf(IsNull(oRsTmp.Fields!TipDocIdentidad), "", oRsTmp.Fields!TipDocIdentidad)
          oRsFox.Fields!NumDocId = IIf(IsNull(oRsTmp.Fields!nroDocumento), "", oRsTmp.Fields!nroDocumento)
          oRsFox.Fields!ResSegur = ""
          oRsFox.Fields!SexoPaci = IIf(IsNull(oRsTmp.Fields!idTipoSexo), "", oRsTmp.Fields!idTipoSexo)
          oRsFox.Fields!FecNacPa = IIf(IsNull(oRsTmp.Fields!FechaNacimiento), "", oRsTmp.Fields!FechaNacimiento)
          oRsFox.Fields!FecIngEm = IIf(IsNull(oRsTmp.Fields!FechaIngreso), "", oRsTmp.Fields!FechaIngreso)
          oRsFox.Fields!HorIngEm = IIf(IsNull(oRsTmp.Fields!HoraIngreso), "", oRsTmp.Fields!HoraIngreso)
          
          With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "Sunasa_TramaPresSaludEmergenciaDx"
             Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, oRsTmp.Fields!idAtencion): .Parameters.Append oParameter
             Set oRsTmp2 = .Execute
             Set oRsTmp2.ActiveConnection = Nothing
          End With
          Set oCommand = Nothing
          Set oParameter = Nothing
          'Inicializar Dx
            lcDx = ""
            lcTipoDx = ""
          If oRsTmp2.RecordCount > 0 Then
            oRsTmp2.MoveFirst
            Do While Not oRsTmp2.EOF
              lcDx = oRsTmp2.Fields!CodigoCie2004
              lcTipoDx = DevuelveCodigoTipoDx(oRsTmp2.Fields!codigo)
              oRsTmp2.MoveNext
            Loop
          End If
          oRsFox.Fields!Diagnost = lcDx
          oRsFox.Fields!TipoDiag = lcTipoDx
          oRsFox.Fields!FecAltaE = IIf(IsNull(oRsTmp.Fields!FechaEgreso), "", oRsTmp.Fields!FechaEgreso)
          oRsFox.Fields!HorAltaE = IIf(IsNull(oRsTmp.Fields!horaEgreso), "", oRsTmp.Fields!horaEgreso)
          oRsFox.Fields!Condicio = ""
          oRsFox.Fields!Resultad = ""
          oRsFox.Update
          
            lcLineaTxtPlano = ""
            lcLineaTxtPlano = lcLineaTxtPlano & Right("00000000000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(280)), 9) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "" & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "" & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!NroHistoriaClinica), "", oRsTmp.Fields!NroHistoriaClinica) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!TipDocIdentidad), "", oRsTmp.Fields!TipDocIdentidad) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!nroDocumento), "", oRsTmp.Fields!nroDocumento) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "" & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!idTipoSexo), "", oRsTmp.Fields!idTipoSexo) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!FechaNacimiento), "", oRsTmp.Fields!FechaNacimiento) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!FechaIngreso), "", oRsTmp.Fields!FechaIngreso) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!HoraIngreso), "", oRsTmp.Fields!HoraIngreso) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & lcDx & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & lcTipoDx & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!FechaEgreso), "", oRsTmp.Fields!FechaEgreso) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!horaEgreso), "", oRsTmp.Fields!horaEgreso) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "" & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "" & "|"
            act.WriteLine (lcLineaTxtPlano)
            
            lnContadorDetalle = lnContadorDetalle + 1
            progresSunasaDetalle.Value = lnContadorDetalle
            Me.Refresh
          
          oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    If oRsFox.State = 1 Then oRsFox.Close
    act.Close
End Sub

Sub GeneraTramaPrestacionesSaludHospitalizacion(oConexion As ADODB.Connection, oConexionFox As ADODB.Connection)
Dim oRsTmp As New Recordset
Dim oRsTmp2 As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRsFox As New Recordset
Dim lcSql As String
Dim lnNumDx As Integer
Dim lcLineaTxtPlano As String
Dim lcPriDxIngreso As String
Dim lcSegDxIngreso As String
Dim lcTerDxIngreso As String
Dim lcPriDxAlta As String
Dim lcSegDxAlta As String
Dim lcTerDxAlta As String
Dim lnContadorDetalle As Long

'Leer datos del SISGalenPlus - PRESTACIONESSALUDHOSPITALIZACION
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "Sunasa_TramaPresSaludHospitalizacion"
        Set oParameter = .CreateParameter("@FechaIni", adDBTimeStamp, adParamInput, 0, Format(Me.txtFechaInicioPresSaludHospitalizacion.Text, "dd/mm/yyyy")): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, Format(Me.txtFechaFinPresSaludHospitalizacion.Text, "dd/mm/yyyy")): .Parameters.Append oParameter
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
   End With
   Set oCommand = Nothing
   Set oParameter = Nothing

   If oRsTmp.RecordCount > 0 Then
        progresSunasaDetalle.Min = 0
        progresSunasaDetalle.Max = oRsTmp.RecordCount
   Else
        progresSunasaDetalle.Min = 0
        progresSunasaDetalle.Max = 1
        progresSunasaDetalle.Value = 1
        Me.Refresh
   End If
   lnContadorDetalle = 0


'Cargar en la tabla temporal Sunasa_PrestacionsSalud_Emergencia
    Dim fso
    Dim act
    Set fso = CreateObject("scripting.filesystemobject")
    lcLineaTxtPlano = ""
    Set act = fso.CreateTextFile(lcBuscaParametro.SeleccionaFilaParametro(313) & "TramaPrestacionesSaludHospitalizacion.txt", True)

    If oRsTmp.RecordCount > 0 Then
       'Inicializa tabla fox
       lcSql = "delete from su_hospi.dbf" 'where Cod_ipre='" & Right("00000000000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(280)), 9) & "'"
       oRsFox.Open "select * from su_hospi.dbf", oConexionFox, adOpenKeyset, adLockOptimistic
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          oRsFox.AddNew
          oRsFox.Fields!Cod_ipre = Right("00000000000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(280)), 9)
          oRsFox.Fields!PeriodoR = ""
          oRsFox.Fields!CodIafas = ""
          oRsFox.Fields!NumHistC = IIf(IsNull(oRsTmp.Fields!NumHistClinica), "", oRsTmp.Fields!NumHistClinica)
          oRsFox.Fields!TipDocId = IIf(IsNull(oRsTmp.Fields!TipDocIdentidad), "", oRsTmp.Fields!TipDocIdentidad)
          oRsFox.Fields!NumDocId = IIf(IsNull(oRsTmp.Fields!NumDocIdentidad), "", oRsTmp.Fields!NumDocIdentidad)
          oRsFox.Fields!RegSegur = ""
          oRsFox.Fields!SexoPaci = IIf(IsNull(oRsTmp.Fields!SexoPaciente), "", oRsTmp.Fields!SexoPaciente)
          oRsFox.Fields!FecNacPa = IIf(IsNull(oRsTmp.Fields!FechaNacimiento), "", oRsTmp.Fields!FechaNacimiento)
          oRsFox.Fields!FecHospi = IIf(IsNull(oRsTmp.Fields!FechaIngreso), "", oRsTmp.Fields!FechaIngreso)
          oRsFox.Fields!AreaProc = IIf(IsNull(oRsTmp.Fields!Proced1), "", oRsTmp.Fields!Proced1)
          oRsFox.Fields!NumColMe = IIf(IsNull(oRsTmp.Fields!mIngNroColegio), "", oRsTmp.Fields!mIngNroColegio)
          
          With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "Sunasa_TramaPresSaludHospitalizacionDxIngreso"
             Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, oRsTmp.Fields!idAtencion): .Parameters.Append oParameter
             Set oRsTmp2 = .Execute
             Set oRsTmp2.ActiveConnection = Nothing
          End With
          Set oCommand = Nothing
          Set oParameter = Nothing
          'Inicializar Dx Ingreso
            lcPriDxIngreso = ""
            lcSegDxIngreso = ""
            lcTerDxIngreso = ""
          lnNumDx = 1
          If oRsTmp2.RecordCount > 0 Then
            oRsTmp2.MoveFirst
            Do While Not oRsTmp2.EOF
              Select Case lnNumDx
              Case 1
                  lcPriDxIngreso = oRsTmp2.Fields!CodigoCie2004
                  lnNumDx = lnNumDx + 1
              Case 2
                  lcSegDxIngreso = oRsTmp2.Fields!CodigoCie2004
                  lnNumDx = lnNumDx + 1
              Case 3
                  lcTerDxIngreso = oRsTmp2.Fields!CodigoCie2004
                  lnNumDx = lnNumDx + 1
              End Select
              oRsTmp2.MoveNext
            Loop
          End If
          oRsFox.Fields!Primerdi = lcPriDxIngreso
          oRsFox.Fields!SegundoD = lcSegDxIngreso
          oRsFox.Fields!TercerDi = lcTerDxIngreso
          oRsFox.Fields!TipPerRe = ""
          oRsFox.Fields!NumColeP = IIf(IsNull(oRsTmp.Fields!mSalNroColegio), "", oRsTmp.Fields!mSalNroColegio)
          
          With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "Sunasa_TramaPresSaludHospitalizacionDxAlta"
             Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, oRsTmp.Fields!idAtencion): .Parameters.Append oParameter
             Set oRsTmp2 = .Execute
             Set oRsTmp2.ActiveConnection = Nothing
          End With
          Set oCommand = Nothing
          Set oParameter = Nothing
          'Inicializar Dx Alta
            lcPriDxAlta = ""
            lcSegDxAlta = ""
            lcTerDxAlta = ""
          lnNumDx = 1
          If oRsTmp2.RecordCount > 0 Then
            oRsTmp2.MoveFirst
            Do While Not oRsTmp2.EOF
              Select Case lnNumDx
              Case 1
                  lcPriDxAlta = oRsTmp2.Fields!CodigoCie2004
                  lnNumDx = lnNumDx + 1
              Case 2
                  lcSegDxAlta = oRsTmp2.Fields!CodigoCie2004
                  lnNumDx = lnNumDx + 1
              Case 3
                  lcTerDxAlta = oRsTmp2.Fields!CodigoCie2004
                  lnNumDx = lnNumDx + 1
              End Select
              oRsTmp2.MoveNext
            Loop
          End If
          oRsFox.Fields!Primerdia = lcPriDxAlta
          oRsFox.Fields!SegundoDi = lcSegDxAlta
          oRsFox.Fields!TercerDia = lcTerDxAlta
          oRsFox.Fields!MotivoAl = ""
          oRsFox.Fields!FecAlta = IIf(IsNull(oRsTmp.Fields!FecAlta), "", oRsTmp.Fields!FecAlta)
          oRsFox.Update
          
            lcLineaTxtPlano = ""
            lcLineaTxtPlano = lcLineaTxtPlano & Right("00000000000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(280)), 9) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "" & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "" & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!NumHistClinica), "", oRsTmp.Fields!NumHistClinica) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!TipDocIdentidad), "", oRsTmp.Fields!TipDocIdentidad) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!NumDocIdentidad), "", oRsTmp.Fields!NumDocIdentidad) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "" & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!SexoPaciente), "", oRsTmp.Fields!SexoPaciente) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!FechaNacimiento), "", oRsTmp.Fields!FechaNacimiento) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!FechaIngreso), "", oRsTmp.Fields!FechaIngreso) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!Proced1), "", oRsTmp.Fields!Proced1) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!mIngNroColegio), "", oRsTmp.Fields!mIngNroColegio) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & lcPriDxAlta & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & lcSegDxAlta & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & lcTerDxAlta & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "" & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!mSalNroColegio), "", oRsTmp.Fields!mSalNroColegio) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & lcPriDxAlta & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & lcSegDxAlta & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & lcTerDxAlta & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "" & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!FecAlta), "", oRsTmp.Fields!FecAlta) & "|"
            act.WriteLine (lcLineaTxtPlano)
            
            lnContadorDetalle = lnContadorDetalle + 1
            progresSunasaDetalle.Value = lnContadorDetalle
            Me.Refresh
          
          
          oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    If oRsFox.State = 1 Then oRsFox.Close
    act.Close
End Sub

Sub GeneraTramaPrestacionesSaludHospParto(oConexion As ADODB.Connection, oConexionFox As ADODB.Connection)
Dim oRsTmp As New Recordset
Dim oRsTmp2 As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRsFox As New Recordset
Dim lcSql As String
Dim lcLineaTxtPlano As String
Dim lcDxAtencion As String
Dim lnContadorDetalle As Long

'Leer datos del SISGalenPlus - PRESTACIONESSALUDHOSPITALIZACION
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "Sunasa_TramaPresSaludHospParto"
        Set oParameter = .CreateParameter("@FechaIni", adDBTimeStamp, adParamInput, 0, Format(Me.txtFechaInicioPresSaludHospParto.Text, "dd/mm/yyyy")): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, Format(Me.txtFechaFinPresSaludHospParto.Text, "dd/mm/yyyy")): .Parameters.Append oParameter
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
   End With
   Set oCommand = Nothing
   Set oParameter = Nothing
   
   
      If oRsTmp.RecordCount > 0 Then
        progresSunasaDetalle.Min = 0
        progresSunasaDetalle.Max = oRsTmp.RecordCount
   Else
        progresSunasaDetalle.Min = 0
        progresSunasaDetalle.Max = 1
        progresSunasaDetalle.Value = 1
        Me.Refresh
   End If
   lnContadorDetalle = 0

'Cargar en la tabla temporal Sunasa_PrestacionsSalud_Emergencia
    Dim fso
    Dim act
    Set fso = CreateObject("scripting.filesystemobject")
    lcLineaTxtPlano = ""
    Set act = fso.CreateTextFile(lcBuscaParametro.SeleccionaFilaParametro(313) & "TramaPrestacionesSaludHospParto.txt", True)
    
    If oRsTmp.RecordCount > 0 Then
       'Inicializa tabla fox
       lcSql = "delete from Su_parto.dbf" 'where Cod_ipre='" & Right("00000000000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(280)), 9) & "'"
       oRsFox.Open "select * from Su_parto.dbf", oConexionFox, adOpenKeyset, adLockOptimistic
       
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          oRsFox.AddNew
          oRsFox.Fields!Cod_ipre = Right("00000000000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(280)), 9)
          oRsFox.Fields!PeriodoR = ""
          oRsFox.Fields!CodIafas = ""
          oRsFox.Fields!NumHisCl = IIf(IsNull(oRsTmp.Fields!NumHistClinica), "", oRsTmp.Fields!NumHistClinica)
          oRsFox.Fields!TipDocId = IIf(IsNull(oRsTmp.Fields!TipDocIdentidad), "", oRsTmp.Fields!TipDocIdentidad)
          oRsFox.Fields!NumDocId = IIf(IsNull(oRsTmp.Fields!NumDocIdentidad), "", oRsTmp.Fields!NumDocIdentidad)
          oRsFox.Fields!RegSegur = ""
          oRsFox.Fields!SexoPaci = IIf(IsNull(oRsTmp.Fields!SexoPaciente), "", oRsTmp.Fields!SexoPaciente)
          oRsFox.Fields!FecNacPa = IIf(IsNull(oRsTmp.Fields!fNacimMadre), "", oRsTmp.Fields!fNacimMadre)
          oRsFox.Fields!FecParto = IIf(IsNull(oRsTmp.Fields!FechaHoraParto), "", oRsTmp.Fields!FechaHoraParto)
          oRsFox.Fields!SemanaGe = IIf(IsNull(oRsTmp.Fields!SemanaGestacion), "", oRsTmp.Fields!SemanaGestacion)
          oRsFox.Fields!TipoPart = ""
          
          With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "TramaPrestacionesSaludHospPartoDx"
             Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, oRsTmp.Fields!idAtencion): .Parameters.Append oParameter
             Set oRsTmp2 = .Execute
             Set oRsTmp2.ActiveConnection = Nothing
          End With
          Set oCommand = Nothing
          Set oParameter = Nothing
          'Inicializar Dx
          lcDxAtencion = ""
          If oRsTmp2.RecordCount > 0 Then
            oRsTmp2.MoveFirst
            Do While Not oRsTmp2.EOF
              lcDxAtencion = oRsTmp2.Fields!CodigoCie2004
              oRsTmp2.MoveNext
            Loop
          End If
          oRsFox.Fields!DiagAten = lcDxAtencion
          oRsFox.Fields!EstadoRe = IIf(IsNull(oRsTmp.Fields!VivoMuerto), "", oRsTmp.Fields!VivoMuerto)
          oRsFox.Fields!HoraNaci = ""
          oRsFox.Fields!PesoNeon = IIf(IsNull(oRsTmp.Fields!PesoNeonato), "", oRsTmp.Fields!PesoNeonato)
          oRsFox.Fields!ApgarUn = IIf(IsNull(oRsTmp.Fields!ApgarPrimerMinuto), "", oRsTmp.Fields!ApgarPrimerMinuto)
          oRsFox.Fields!ApgarCin = IIf(IsNull(oRsTmp.Fields!ApgarCincoMinutos), "", oRsTmp.Fields!ApgarCincoMinutos)
          oRsFox.Fields!FechaCir = ""
          oRsFox.Update
          
            lcLineaTxtPlano = ""
            lcLineaTxtPlano = lcLineaTxtPlano & Right("00000000000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(280)), 9) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "" & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "" & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!NumHistClinica), "", oRsTmp.Fields!NumHistClinica) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!TipDocIdentidad), "", oRsTmp.Fields!TipDocIdentidad) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!NumDocIdentidad), "", oRsTmp.Fields!NumDocIdentidad) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "" & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!SexoPaciente), "", oRsTmp.Fields!SexoPaciente) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!fNacimMadre), "", oRsTmp.Fields!fNacimMadre) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!FechaHoraParto), "", oRsTmp.Fields!FechaHoraParto) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!SemanaGestacion), "", oRsTmp.Fields!SemanaGestacion) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "" & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & lcDxAtencion & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!VivoMuerto), "", oRsTmp.Fields!VivoMuerto) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "" & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!PesoNeonato), "", oRsTmp.Fields!PesoNeonato) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!ApgarPrimerMinuto), "", oRsTmp.Fields!ApgarPrimerMinuto) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & IIf(IsNull(oRsTmp.Fields!ApgarCincoMinutos), "", oRsTmp.Fields!ApgarCincoMinutos) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & "" & "|"
            act.WriteLine (lcLineaTxtPlano)
            
            lnContadorDetalle = lnContadorDetalle + 1
            progresSunasaDetalle.Value = lnContadorDetalle
            Me.Refresh
        
          oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    If oRsFox.State = 1 Then oRsFox.Close
    act.Close
End Sub

Public Function DevuelveCodigoTipoDx(codigo As String) As String
    DevuelveCodigoTipoDx = ""
    Select Case codigo
    Case "D"
    DevuelveCodigoTipoDx = "02"
    Case "P"
    DevuelveCodigoTipoDx = "01"
    Case "R"
    DevuelveCodigoTipoDx = "03"
    End Select
End Function

Sub MostrarFormulario()
Me.Show 1
End Sub



Private Sub Form_Load()
    txtFProgInicio.Text = sighentidades.PrimerFechaDDMMYYDelMesActual()
    txtFprogFinal.Text = Date
    txtFreferInicio.Text = sighentidades.PrimerFechaDDMMYYDelMesActual()
    txtFreferFin.Text = Date
    txtfechaInicialEmisionCitas.Text = sighentidades.PrimerFechaDDMMYYDelMesActual()
    txtFechaFinEmisionCitas.Text = Date
    txtFechaInicioPresSaludEmergencia.Text = sighentidades.PrimerFechaDDMMYYDelMesActual()
    txtFechaFinPresSaludEmergencia.Text = Date
    txtFechaInicioPresSaludHospitalizacion.Text = sighentidades.PrimerFechaDDMMYYDelMesActual()
    txtFechaFinPresSaludHospitalizacion.Text = Date
    txtFechaInicioPresSaludHospParto.Text = sighentidades.PrimerFechaDDMMYYDelMesActual()
    txtFechaFinPresSaludHospParto.Text = Date
    txtFcptInicio.Text = sighentidades.PrimerFechaDDMMYYDelMesActual()
    txtFcptFin.Text = Date
    txtFrecurInicial.Text = sighentidades.PrimerFechaDDMMYYDelMesActual()
    txtFrecurFinal.Text = Date
End Sub


Sub GeneraTramaPrestacionesSaludHospParto2016(oConexion As ADODB.Connection, lcIpress As String, lcUgipress As String, _
                                              lcParametro313 As String)
Dim oRsTmp As New Recordset
Dim oRsTmp2 As New Recordset
Dim oRsTabla As New Recordset
Dim oRsFox As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim lcSql As String
Dim lcLineaTxtPlano As String
Dim lcDxAtencion As String
Dim lnContadorDetalle As Long, lcPeriodo As String, lnIdAtencion As Long
Dim lnVivos As Long, lnMuertos As Long, lcTipoParto As String, lcComplicacion As String, lnNuevo As Boolean
Const lcConComplicacion As String = "/O60/O61/O62/O75/O81/O83/O84/"
Const lcPartoCesaria As String = "/O82.0/O82.1/O82.2/O82.8/O82.9/084.2/"
    lblTabla.Caption = chbPresSaludHospParto.Caption
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "Sunasa_TramaPresSaludHospParto"
        Set oParameter = .CreateParameter("@FechaIni", adDBTimeStamp, adParamInput, 0, Format(Me.txtFechaInicioPresSaludHospParto.Text, "dd/mm/yyyy")): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, Format(Me.txtFechaFinPresSaludHospParto.Text, "dd/mm/yyyy")): .Parameters.Append oParameter
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
   End With
   Set oCommand = Nothing
   Set oParameter = Nothing
   If oRsTmp.RecordCount > 0 Then
        With oRsTabla
              .Fields.Append "TipoComplicacion", adVarChar, 4, adFldIsNullable
              .Fields.Append "partos", adInteger
              .Fields.Append "nacimientos", adInteger
              .Fields.Append "nacVivos", adInteger
              .Fields.Append "nacMuertos", adInteger
              .LockType = adLockOptimistic
              .Open
              .AddNew
              .Fields!TipoComplicacion = "0101"
              .Fields!partos = 0
              .Fields!nacimientos = 0
              .Fields!nacVivos = 0
              .Fields!nacMuertos = 0
              .Update
              .AddNew
              .Fields!TipoComplicacion = "0102"
              .Fields!partos = 0
              .Fields!nacimientos = 0
              .Fields!nacVivos = 0
              .Fields!nacMuertos = 0
              .Update
              .AddNew
              .Fields!TipoComplicacion = "0201"
              .Fields!partos = 0
              .Fields!nacimientos = 0
              .Fields!nacVivos = 0
              .Fields!nacMuertos = 0
              .Update
              .AddNew
              .Fields!TipoComplicacion = "0202"
              .Fields!partos = 0
              .Fields!nacimientos = 0
              .Fields!nacVivos = 0
              .Fields!nacMuertos = 0
              .Update
        End With
        progresSunasa.Min = 0
        progresSunasa.Max = oRsTmp.RecordCount
        lnContadorDetalle = 0
        oRsTmp.MoveFirst
        Do While Not oRsTmp.EOF
           lnVivos = 0: lnMuertos = 0
           lnIdAtencion = oRsTmp!idAtencion
           Do While Not oRsTmp.EOF And lnIdAtencion = oRsTmp!idAtencion
                If oRsTmp!VivoMuerto = 1 Then
                   lnVivos = lnVivos + 1
                Else
                   lnMuertos = lnMuertos + 1
                End If
                oRsTmp.MoveNext
                If oRsTmp.EOF Then
                   Exit Do
                End If
           Loop
           lcTipoParto = "01": lcComplicacion = "01"
           Set oRsTmp2 = mo_ReglasAdmision.AtencionesDiagnosticosSeleccionarXidAtencion(lnIdAtencion, oConexion)
           If oRsTmp2.RecordCount > 0 Then
                oRsTmp2.MoveFirst
                Do While Not oRsTmp2.EOF
                   If InStr(oRsTmp2!CodigoCie2004, lcPartoCesaria) > 0 Then   'Cesárea
                      lcTipoParto = "02"
                   End If
                   If oRsTmp2!IdClasificacionDx = 6 Or _
                                             InStr(Left(oRsTmp2!CodigoCie2004, 3), lcConComplicacion) > 0 Then                                 'hubo complicacion
                      lcComplicacion = "02"
                   End If
                   oRsTmp2.MoveNext
                Loop
            End If
            oRsTmp2.Close
            lnNuevo = True
            If oRsTabla.RecordCount > 0 Then
               oRsTabla.MoveFirst
               oRsTabla.Find "TipoComplicacion='" & lcTipoParto & lcComplicacion & "'"
               If Not oRsTabla.EOF Then
                  lnNuevo = False
               End If
            End If
            If lnNuevo = True Then
               oRsTabla.AddNew
               oRsTabla.Fields!TipoComplicacion = lcTipoParto & lcComplicacion
            End If
            oRsTabla.Fields!partos = oRsTabla.Fields!partos + 1
            oRsTabla.Fields!nacimientos = oRsTabla.Fields!nacimientos + lnMuertos + lnVivos
            oRsTabla.Fields!nacVivos = oRsTabla.Fields!nacVivos + lnVivos
            oRsTabla.Fields!nacMuertos = oRsTabla.Fields!nacMuertos + lnMuertos
            oRsTabla.Update
            
            lnContadorDetalle = lnContadorDetalle + 1
            DoEvents
            progresSunasa.Value = lnContadorDetalle
            Me.Refresh
            
        Loop
        Dim fso
        Dim act
        Set fso = CreateObject("scripting.filesystemobject")
        lcLineaTxtPlano = lcIpress & "_" & Right(txtFechaInicioPresSaludHospParto.Text, 4) & "_" & _
                          Mid(txtFechaInicioPresSaludHospParto.Text, 4, 2) & "_TAE0.TXT"
        Set act = fso.CreateTextFile(lcParametro313 & lcLineaTxtPlano, True)
        progresSunasaDetalle.Min = 0
        progresSunasaDetalle.Max = oRsTabla.RecordCount
        lnContadorDetalle = 0
        lcPeriodo = Right(txtFechaInicioPresSaludHospParto.Text, 4) & Mid(txtFechaInicioPresSaludHospParto.Text, 4, 2)
        oRsTabla.MoveFirst
        Do While Not oRsTabla.EOF
            lcLineaTxtPlano = ""
            lcLineaTxtPlano = lcLineaTxtPlano & lcPeriodo & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & lcIpress & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & lcUgipress & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & Left(oRsTabla!TipoComplicacion, 2) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & Right(oRsTabla!TipoComplicacion, 2) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & Trim(str(oRsTabla!partos)) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & Trim(str(oRsTabla!nacimientos)) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & Trim(str(oRsTabla!nacVivos)) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & Trim(str(oRsTabla!nacMuertos))
            act.WriteLine (lcLineaTxtPlano)
            
            lnContadorDetalle = lnContadorDetalle + 1
            DoEvents
            progresSunasaDetalle.Value = lnContadorDetalle
            Me.Refresh
            
            oRsTabla.MoveNext
        Loop
        act.Close
   Else
        progresSunasaDetalle.Min = 0
        progresSunasaDetalle.Max = 1
        progresSunasaDetalle.Value = 1
        Me.Refresh
   End If
    Set oRsTmp = Nothing
    Set oRsTmp2 = Nothing
    Set oRsTabla = Nothing
    Set oRsFox = Nothing

End Sub


Sub GeneraTramaEmisionCitas2016(oConexion As ADODB.Connection, lcIpress As String, lcUgipress As String, _
                                              lcParametro313 As String)

Dim rsReporte As New Recordset
Dim oRsTmp2 As New Recordset
Dim oRsTabla1 As New Recordset
Dim oRsTabla11 As New Recordset
Dim oRsTabla2 As New Recordset
Dim oRsTabla22 As New Recordset
Dim oRsFox As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim lcSql As String
Dim lcLineaTxtPlano As String
Dim lcDxAtencion As String
Dim lnContadorDetalle As Long, lcPeriodo As String, lnIdAtencion As Long, lnNuevo As Boolean, lnIdTipoServicio As Integer
Dim lcSexo As String, lcGrupo As String, lcDx As String, lcColegio As String, mda_FechaInicio As Date, mda_FechaFin As Date
Dim lbPersonaA As Boolean, lbPersonaDx As Boolean
        lblTabla.Caption = chbEmisionCitas.Caption
        mda_FechaInicio = CDate(txtfechaInicialEmisionCitas.Text)
        mda_FechaFin = CDate(txtFechaFinEmisionCitas.Text)
        lnIdTipoServicio = sghTipoServicio.sghConsultaExterna
        Set rsReporte = mo_AdminReportes.ReporteEgresosHospitalarios(0, 0, 0, mda_FechaInicio, mda_FechaFin, _
                                                                            lnIdTipoServicio)
        If rsReporte.RecordCount > 0 Then
           With oRsTabla1
                  .Fields.Append "sexo", adVarChar, 1
                  .Fields.Append "grupo", adVarChar, 2
                  .Fields.Append "atencionesM", adInteger
                  .Fields.Append "atencionesNOM", adInteger
                  .Fields.Append "atendidos", adInteger
                  .LockType = adLockOptimistic
                  .Open
           End With
           With oRsTabla11
                  .Fields.Append "sexo", adVarChar, 1
                  .Fields.Append "grupo", adVarChar, 2
                  .Fields.Append "NroHistoriaClinica", adInteger
                  .Fields.Append "TipoNumeracion", adVarChar, 100
                  .LockType = adLockOptimistic
                  .Open
           End With
           With oRsTabla2
                  .Fields.Append "sexo", adVarChar, 1
                  .Fields.Append "grupo", adVarChar, 2
                  .Fields.Append "dx", adVarChar, 5
                  .Fields.Append "atendidos", adInteger
                  .LockType = adLockOptimistic
                  .Open
           End With
           With oRsTabla22
                  .Fields.Append "sexo", adVarChar, 1
                  .Fields.Append "grupo", adVarChar, 2
                  .Fields.Append "dx", adVarChar, 5
                  .Fields.Append "NroHistoriaClinica", adInteger
                  .Fields.Append "TipoNumeracion", adVarChar, 100
                  .LockType = adLockOptimistic
                  .Open
           End With
           progresSunasaDetalle.Min = 0
           progresSunasaDetalle.Max = rsReporte.RecordCount + 1
           lnContadorDetalle = 0
           
           rsReporte.MoveFirst
           Do While Not rsReporte.EOF
              lcSexo = IIf(UCase(Left(rsReporte!Sexo, 1)) = "M", "1", "2")
              lcGrupo = ""
              If rsReporte!TipoEdad = "A" Then
                 If rsReporte!Edad >= 1 And rsReporte!Edad <= 4 Then
                    lcGrupo = "2"
                 ElseIf rsReporte!Edad >= 5 And rsReporte!Edad <= 9 Then
                    lcGrupo = "3"
                 ElseIf rsReporte!Edad >= 10 And rsReporte!Edad <= 14 Then
                    lcGrupo = "4"
                 ElseIf rsReporte!Edad >= 15 And rsReporte!Edad <= 19 Then
                    lcGrupo = "5"
                 ElseIf rsReporte!Edad >= 20 And rsReporte!Edad <= 24 Then
                    lcGrupo = "6"
                 ElseIf rsReporte!Edad >= 25 And rsReporte!Edad <= 29 Then
                    lcGrupo = "7"
                 ElseIf rsReporte!Edad >= 30 And rsReporte!Edad <= 34 Then
                    lcGrupo = "8"
                 ElseIf rsReporte!Edad >= 35 And rsReporte!Edad <= 39 Then
                    lcGrupo = "9"
                 ElseIf rsReporte!Edad >= 40 And rsReporte!Edad <= 44 Then
                    lcGrupo = "10"
                 ElseIf rsReporte!Edad >= 45 And rsReporte!Edad <= 49 Then
                    lcGrupo = "11"
                 ElseIf rsReporte!Edad >= 50 And rsReporte!Edad <= 54 Then
                    lcGrupo = "12"
                 ElseIf rsReporte!Edad >= 55 And rsReporte!Edad <= 59 Then
                    lcGrupo = "13"
                 ElseIf rsReporte!Edad >= 60 And rsReporte!Edad <= 64 Then
                    lcGrupo = "14"
                 Else
                    lcGrupo = "15"
                 End If
              Else
                 lcGrupo = "1"
              End If
              '
              lcDx = ""
              If lnIdTipoServicio <> sghTipoServicio.sghConsultaExterna Then
                    Set oRsTmp2 = mo_AdminReportes.ReporteAtencionesDiagnosticosDeEgreso(rsReporte!idAtencion)
              Else
                    Set oRsTmp2 = mo_ReglasAdmision.BuscaCEAtencionesDx(rsReporte!idAtencion)
              End If
              oRsTmp2.Filter = "codigoDx<>null"
              If oRsTmp2.RecordCount > 0 Then
                   lcDx = Left(oRsTmp2!CodigoDx, 5)
              End If
              oRsTmp2.Close
              '
              lcColegio = ""
              If lnIdTipoServicio = sghTipoServicio.sghConsultaExterna Then
                    Set oRsTmp2 = mo_AdminServiciosComunes.AtencionesSeleccionarMedicoPorCuenta(rsReporte!idCuentaAtencion)
              Else
                    Set oRsTmp2 = mo_AdminServiciosComunes.AtencionesSeleccionarMedicoEgresoPorCuenta(rsReporte!idCuentaAtencion)
              End If
              If oRsTmp2.RecordCount > 0 Then
                 lcColegio = oRsTmp2!idColegioHIS
              End If
              oRsTmp2.Close
              'Personas Asistencial
              lbPersonaA = True
              If oRsTabla11.RecordCount > 0 Then
                 oRsTabla11.MoveFirst
                 Do While Not oRsTabla11.EOF
                    If oRsTabla11!Sexo = lcSexo And oRsTabla11!Grupo = lcGrupo And _
                                                    oRsTabla11!NroHistoriaClinica = rsReporte!NroHistoriaClinica And _
                                                    oRsTabla11!TipoNumeracion = rsReporte!TipoNumeracion Then
                       lbPersonaA = False
                       Exit Do
                    End If
                    oRsTabla11.MoveNext
                 Loop
              End If
              If lbPersonaA = True Then
                 oRsTabla11.AddNew
                 oRsTabla11.Fields!Sexo = lcSexo
                 oRsTabla11.Fields!Grupo = lcGrupo
                 oRsTabla11.Fields!NroHistoriaClinica = rsReporte!NroHistoriaClinica
                 oRsTabla11.Fields!TipoNumeracion = rsReporte!TipoNumeracion
                 oRsTabla11.Update
              End If
              'asistencial
              lnNuevo = True
              If oRsTabla1.RecordCount > 0 Then
                 oRsTabla1.MoveFirst
                 Do While Not oRsTabla1.EOF
                    If oRsTabla1!Sexo = lcSexo And oRsTabla1!Grupo = lcGrupo Then
                       lnNuevo = False
                       Exit Do
                    End If
                    oRsTabla1.MoveNext
                 Loop
              End If
              If lnNuevo = True Then
                    oRsTabla1.AddNew
                    oRsTabla1.Fields!Sexo = lcSexo
                    oRsTabla1.Fields!Grupo = lcGrupo
                    If lcColegio = "01" Then
                       oRsTabla1.Fields!atencionesM = 1
                       oRsTabla1.Fields!atencionesNOM = 0
                    Else
                       oRsTabla1.Fields!atencionesM = 0
                       oRsTabla1.Fields!atencionesNOM = 1
                    End If
                    If lbPersonaA = True Then
                       oRsTabla1.Fields!Atendidos = 1
                    End If
              Else
                    If lcColegio = "01" Then
                       oRsTabla1.Fields!atencionesM = oRsTabla1.Fields!atencionesM + 1
                    Else
                       oRsTabla1.Fields!atencionesNOM = oRsTabla1.Fields!atencionesNOM + 1
                    End If
                    If lbPersonaA = True Then
                       oRsTabla1.Fields!Atendidos = oRsTabla1.Fields!Atendidos + 1
                    End If
              End If
              oRsTabla1.Update
              'Personas dx
              If lcDx <> "" Then
                    lbPersonaDx = True
                    If oRsTabla22.RecordCount > 0 Then
                       oRsTabla22.MoveFirst
                       Do While Not oRsTabla22.EOF
                          If oRsTabla22!Sexo = lcSexo And oRsTabla22!Grupo = lcGrupo And oRsTabla22!dx = lcDx And _
                                                          oRsTabla22!NroHistoriaClinica = rsReporte!NroHistoriaClinica And _
                                                          oRsTabla22!TipoNumeracion = rsReporte!TipoNumeracion Then
                             lbPersonaDx = False
                             Exit Do
                          End If
                          oRsTabla22.MoveNext
                       Loop
                    End If
                    If lbPersonaDx = True Then
                       oRsTabla22.AddNew
                       oRsTabla22.Fields!Sexo = lcSexo
                       oRsTabla22.Fields!Grupo = lcGrupo
                       oRsTabla22.Fields!dx = lcDx
                       oRsTabla22.Fields!NroHistoriaClinica = rsReporte!NroHistoriaClinica
                       oRsTabla22.Fields!TipoNumeracion = rsReporte!TipoNumeracion
                       oRsTabla22.Update
                    End If
                    'dx
                    lnNuevo = True
                    If oRsTabla2.RecordCount > 0 Then
                       oRsTabla2.MoveFirst
                       Do While Not oRsTabla2.EOF
                          If oRsTabla2!Sexo = lcSexo And oRsTabla2!Grupo = lcGrupo And oRsTabla2.Fields!dx = lcDx Then
                             lnNuevo = False
                             Exit Do
                          End If
                          oRsTabla2.MoveNext
                       Loop
                    End If
                    If lnNuevo = True Then
                          oRsTabla2.AddNew
                          oRsTabla2.Fields!Sexo = lcSexo
                          oRsTabla2.Fields!Grupo = lcGrupo
                          oRsTabla2.Fields!dx = lcDx
                          If lbPersonaDx = True Then
                             oRsTabla2.Fields!Atendidos = 1
                          End If
                    Else
                          If lbPersonaDx = True Then
                             oRsTabla2.Fields!Atendidos = oRsTabla2.Fields!Atendidos + 1
                          End If
                    End If
                    oRsTabla2.Update
              End If
              '
              DoEvents
              progresSunasaDetalle.Value = lnContadorDetalle
              Me.Refresh
              lnContadorDetalle = lnContadorDetalle + 1
            
              rsReporte.MoveNext
           Loop


            Dim fso
            Dim act
            Set fso = CreateObject("scripting.filesystemobject")
            '
            lcLineaTxtPlano = lcIpress & "_" & Right(txtfechaInicialEmisionCitas.Text, 4) & "_" & _
                              Mid(txtfechaInicialEmisionCitas.Text, 4, 2) & "_TAB1.TXT"
            Set act = fso.CreateTextFile(lcParametro313 & lcLineaTxtPlano, True)
            progresSunasaDetalle.Min = 0
            progresSunasaDetalle.Max = oRsTabla1.RecordCount
            lnContadorDetalle = 0
            lcPeriodo = Right(txtfechaInicialEmisionCitas.Text, 4) & Mid(txtfechaInicialEmisionCitas.Text, 4, 2)
            If oRsTabla1.RecordCount > 0 Then
                oRsTabla1.MoveFirst
                Do While Not oRsTabla1.EOF
                    lcLineaTxtPlano = ""
                    lcLineaTxtPlano = lcLineaTxtPlano & lcPeriodo & "|"
                    lcLineaTxtPlano = lcLineaTxtPlano & lcIpress & "|"
                    lcLineaTxtPlano = lcLineaTxtPlano & lcUgipress & "|"
                    lcLineaTxtPlano = lcLineaTxtPlano & oRsTabla1!Sexo & "|"
                    lcLineaTxtPlano = lcLineaTxtPlano & oRsTabla1!Grupo & "|"
                    lcLineaTxtPlano = lcLineaTxtPlano & Trim(str(oRsTabla1!atencionesM)) & "|"
                    lcLineaTxtPlano = lcLineaTxtPlano & Trim(str(oRsTabla1!atencionesNOM)) & "|"
                    lcLineaTxtPlano = lcLineaTxtPlano & Trim(str(oRsTabla1!Atendidos))
                    act.WriteLine (lcLineaTxtPlano)
                    
                    lnContadorDetalle = lnContadorDetalle + 1
                    DoEvents
                    progresSunasaDetalle.Value = lnContadorDetalle
                    Me.Refresh
                    
                    oRsTabla1.MoveNext
                Loop
            End If
            act.Close
            '
            
            lcLineaTxtPlano = lcIpress & "_" & Right(txtfechaInicialEmisionCitas.Text, 4) & "_" & _
                              Mid(txtfechaInicialEmisionCitas.Text, 4, 2) & "_TAB2.TXT"
            Set act = fso.CreateTextFile(lcParametro313 & lcLineaTxtPlano, True)
            progresSunasaDetalle.Min = 0
            progresSunasaDetalle.Max = oRsTabla2.RecordCount
            lnContadorDetalle = 0
            lcPeriodo = Right(txtfechaInicialEmisionCitas.Text, 4) & Mid(txtfechaInicialEmisionCitas.Text, 4, 2)
            If oRsTabla2.RecordCount > 0 Then
                oRsTabla2.MoveFirst
                Do While Not oRsTabla2.EOF
                    lcLineaTxtPlano = ""
                    lcLineaTxtPlano = lcLineaTxtPlano & lcPeriodo & "|"
                    lcLineaTxtPlano = lcLineaTxtPlano & lcIpress & "|"
                    lcLineaTxtPlano = lcLineaTxtPlano & lcUgipress & "|"
                    lcLineaTxtPlano = lcLineaTxtPlano & oRsTabla2!Sexo & "|"
                    lcLineaTxtPlano = lcLineaTxtPlano & oRsTabla2!Grupo & "|"
                    lcLineaTxtPlano = lcLineaTxtPlano & oRsTabla2!dx & "|"
                    lcLineaTxtPlano = lcLineaTxtPlano & Trim(str(oRsTabla2!Atendidos))
                    act.WriteLine (lcLineaTxtPlano)
                    
                    lnContadorDetalle = lnContadorDetalle + 1
                    DoEvents
                    progresSunasaDetalle.Value = lnContadorDetalle
                    Me.Refresh
                    
                    oRsTabla2.MoveNext
                Loop
            End If
            act.Close
        
        
    Else
            progresSunasaDetalle.Min = 0
            progresSunasaDetalle.Max = 1
            progresSunasaDetalle.Value = 1
            Me.Refresh
    End If
    rsReporte.Close
    Set rsReporte = Nothing
    Set oRsTmp2 = Nothing
    Set oRsTabla1 = Nothing
    Set oRsTabla2 = Nothing
    Set oRsTabla11 = Nothing
    Set oRsTabla22 = Nothing
    Set oRsFox = Nothing
End Sub

Sub GeneraTramaEmisionEmergencia2016(oConexion As ADODB.Connection, lcIpress As String, lcUgipress As String, _
                                              lcParametro313 As String)

Dim rsReporte As New Recordset
Dim oRsTmp2 As New Recordset
Dim oRsTabla1 As New Recordset
Dim oRsTabla11 As New Recordset
Dim oRsTabla2 As New Recordset
Dim oRsTabla22 As New Recordset
Dim oRsFox As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim lcSql As String
Dim lcLineaTxtPlano As String
Dim lcDxAtencion As String
Dim lnContadorDetalle As Long, lcPeriodo As String, lnIdAtencion As Long, lnNuevo As Boolean, lnIdTipoServicio As Integer
Dim lcSexo As String, lcGrupo As String, lcDx As String, lcColegio As String, mda_FechaInicio As Date, mda_FechaFin As Date
Dim lbPersonaA As Boolean, lbPersonaDx As Boolean
        lblTabla.Caption = chbPresSaludEmergencia.Caption
        mda_FechaInicio = CDate(txtFechaInicioPresSaludEmergencia.Text)
        mda_FechaFin = CDate(txtFechaFinPresSaludEmergencia.Text)
        lnIdTipoServicio = sghTipoServicio.sghEmergenciaConsultorios
        Set rsReporte = mo_AdminReportes.ReporteEgresosHospitalarios(0, 0, 0, mda_FechaInicio, mda_FechaFin, _
                                                                            lnIdTipoServicio)
        If rsReporte.RecordCount > 0 Then
           With oRsTabla1
                  .Fields.Append "sexo", adVarChar, 1
                  .Fields.Append "grupo", adVarChar, 2
                  .Fields.Append "atenciones", adInteger
                  .Fields.Append "atendidos", adInteger
                  .LockType = adLockOptimistic
                  .Open
           End With
           With oRsTabla11
                  .Fields.Append "sexo", adVarChar, 1
                  .Fields.Append "grupo", adVarChar, 2
                  .Fields.Append "NroHistoriaClinica", adInteger
                  .Fields.Append "TipoNumeracion", adVarChar, 100
                  .LockType = adLockOptimistic
                  .Open
           End With
           With oRsTabla2
                  .Fields.Append "sexo", adVarChar, 1
                  .Fields.Append "grupo", adVarChar, 2
                  .Fields.Append "dx", adVarChar, 5
                  .Fields.Append "atendidos", adInteger
                  .LockType = adLockOptimistic
                  .Open
           End With
           With oRsTabla22
                  .Fields.Append "sexo", adVarChar, 1
                  .Fields.Append "grupo", adVarChar, 2
                  .Fields.Append "dx", adVarChar, 5
                  .Fields.Append "NroHistoriaClinica", adInteger
                  .Fields.Append "TipoNumeracion", adVarChar, 100
                  .LockType = adLockOptimistic
                  .Open
           End With
           progresSunasaDetalle.Min = 0
           progresSunasaDetalle.Max = rsReporte.RecordCount
           lnContadorDetalle = 0
           
           rsReporte.MoveFirst
           Do While Not rsReporte.EOF
              lcSexo = IIf(UCase(Left(rsReporte!Sexo, 1)) = "M", "1", "2")
              lcGrupo = ""
              If rsReporte!TipoEdad = "A" Then
                 If rsReporte!Edad >= 1 And rsReporte!Edad <= 4 Then
                    lcGrupo = "2"
                 ElseIf rsReporte!Edad >= 5 And rsReporte!Edad <= 9 Then
                    lcGrupo = "3"
                 ElseIf rsReporte!Edad >= 10 And rsReporte!Edad <= 14 Then
                    lcGrupo = "4"
                 ElseIf rsReporte!Edad >= 15 And rsReporte!Edad <= 19 Then
                    lcGrupo = "5"
                 ElseIf rsReporte!Edad >= 20 And rsReporte!Edad <= 24 Then
                    lcGrupo = "6"
                 ElseIf rsReporte!Edad >= 25 And rsReporte!Edad <= 29 Then
                    lcGrupo = "7"
                 ElseIf rsReporte!Edad >= 30 And rsReporte!Edad <= 34 Then
                    lcGrupo = "8"
                 ElseIf rsReporte!Edad >= 35 And rsReporte!Edad <= 39 Then
                    lcGrupo = "9"
                 ElseIf rsReporte!Edad >= 40 And rsReporte!Edad <= 44 Then
                    lcGrupo = "10"
                 ElseIf rsReporte!Edad >= 45 And rsReporte!Edad <= 49 Then
                    lcGrupo = "11"
                 ElseIf rsReporte!Edad >= 50 And rsReporte!Edad <= 54 Then
                    lcGrupo = "12"
                 ElseIf rsReporte!Edad >= 55 And rsReporte!Edad <= 59 Then
                    lcGrupo = "13"
                 ElseIf rsReporte!Edad >= 60 And rsReporte!Edad <= 64 Then
                    lcGrupo = "14"
                 Else
                    lcGrupo = "15"
                 End If
              Else
                 lcGrupo = "1"
              End If
              '
              lcDx = ""
              If lnIdTipoServicio <> sghTipoServicio.sghConsultaExterna Then
                    Set oRsTmp2 = mo_AdminReportes.ReporteAtencionesDiagnosticosDeEgreso(rsReporte!idAtencion)
              Else
                    Set oRsTmp2 = mo_ReglasAdmision.BuscaCEAtencionesDx(rsReporte!idAtencion)
              End If
              oRsTmp2.Filter = "codigoDx<>null"
              If oRsTmp2.RecordCount > 0 Then
                   lcDx = Left(oRsTmp2!CodigoDx, 5)
              End If
              oRsTmp2.Close

              'Personas Asistencial
              lbPersonaA = True
              If oRsTabla11.RecordCount > 0 Then
                 oRsTabla11.MoveFirst
                 Do While Not oRsTabla11.EOF
                    If oRsTabla11!Sexo = lcSexo And oRsTabla11!Grupo = lcGrupo And _
                                                    oRsTabla11!NroHistoriaClinica = rsReporte!NroHistoriaClinica And _
                                                    oRsTabla11!TipoNumeracion = rsReporte!TipoNumeracion Then
                       lbPersonaA = False
                       Exit Do
                    End If
                    oRsTabla11.MoveNext
                 Loop
              End If
              If lbPersonaA = True Then
                 oRsTabla11.AddNew
                 oRsTabla11.Fields!Sexo = lcSexo
                 oRsTabla11.Fields!Grupo = lcGrupo
                 oRsTabla11.Fields!NroHistoriaClinica = rsReporte!NroHistoriaClinica
                 oRsTabla11.Fields!TipoNumeracion = rsReporte!TipoNumeracion
                 oRsTabla11.Update
              End If
              'asistencial
              lnNuevo = True
              If oRsTabla1.RecordCount > 0 Then
                 oRsTabla1.MoveFirst
                 Do While Not oRsTabla1.EOF
                    If oRsTabla1!Sexo = lcSexo And oRsTabla1!Grupo = lcGrupo Then
                       lnNuevo = False
                       Exit Do
                    End If
                    oRsTabla1.MoveNext
                 Loop
              End If
              If lnNuevo = True Then
                    oRsTabla1.AddNew
                    oRsTabla1.Fields!Sexo = lcSexo
                    oRsTabla1.Fields!Grupo = lcGrupo
                    oRsTabla1.Fields!Atenciones = 1
                    If lbPersonaA = True Then
                       oRsTabla1.Fields!Atendidos = 1
                    End If
              Else
                    oRsTabla1.Fields!Atenciones = oRsTabla1.Fields!Atenciones + 1
                    If lbPersonaA = True Then
                       oRsTabla1.Fields!Atendidos = oRsTabla1.Fields!Atendidos + 1
                    End If
              End If
              oRsTabla1.Update
              'Personas dx
              If lcDx <> "" Then
                    lbPersonaDx = True
                    If oRsTabla22.RecordCount > 0 Then
                       oRsTabla22.MoveFirst
                       Do While Not oRsTabla22.EOF
                          If oRsTabla22!Sexo = lcSexo And oRsTabla22!Grupo = lcGrupo And oRsTabla22!dx = lcDx And _
                                                          oRsTabla22!NroHistoriaClinica = rsReporte!NroHistoriaClinica And _
                                                          oRsTabla22!TipoNumeracion = rsReporte!TipoNumeracion Then
                             lbPersonaDx = False
                             Exit Do
                          End If
                          oRsTabla22.MoveNext
                       Loop
                    End If
                    If lbPersonaDx = True Then
                       oRsTabla22.AddNew
                       oRsTabla22.Fields!Sexo = lcSexo
                       oRsTabla22.Fields!Grupo = lcGrupo
                       oRsTabla22.Fields!dx = lcDx
                       oRsTabla22.Fields!NroHistoriaClinica = rsReporte!NroHistoriaClinica
                       oRsTabla22.Fields!TipoNumeracion = rsReporte!TipoNumeracion
                       oRsTabla22.Update
                    End If
                    'dx
                    lnNuevo = True
                    If oRsTabla2.RecordCount > 0 Then
                       oRsTabla2.MoveFirst
                       Do While Not oRsTabla2.EOF
                          If oRsTabla2!Sexo = lcSexo And oRsTabla2!Grupo = lcGrupo And oRsTabla2.Fields!dx = lcDx Then
                             lnNuevo = False
                             Exit Do
                          End If
                          oRsTabla2.MoveNext
                       Loop
                    End If
                    If lnNuevo = True Then
                          oRsTabla2.AddNew
                          oRsTabla2.Fields!Sexo = lcSexo
                          oRsTabla2.Fields!Grupo = lcGrupo
                          oRsTabla2.Fields!dx = lcDx
                          If lbPersonaDx = True Then
                             oRsTabla2.Fields!Atendidos = 1
                          End If
                    Else
                          If lbPersonaDx = True Then
                             oRsTabla2.Fields!Atendidos = oRsTabla2.Fields!Atendidos + 1
                          End If
                    End If
                    oRsTabla2.Update
              End If
              '
              DoEvents
              progresSunasaDetalle.Value = lnContadorDetalle
              Me.Refresh
              lnContadorDetalle = lnContadorDetalle + 1
              
              rsReporte.MoveNext
           Loop


            Dim fso
            Dim act
            Set fso = CreateObject("scripting.filesystemobject")
            '
            lcLineaTxtPlano = lcIpress & "_" & Right(txtFechaInicioPresSaludEmergencia.Text, 4) & "_" & _
                              Mid(txtFechaInicioPresSaludEmergencia.Text, 4, 2) & "_TAC1.TXT"
            Set act = fso.CreateTextFile(lcParametro313 & lcLineaTxtPlano, True)
            progresSunasaDetalle.Min = 0
            progresSunasaDetalle.Max = oRsTabla1.RecordCount
            lnContadorDetalle = 0
            lcPeriodo = Right(txtFechaInicioPresSaludEmergencia.Text, 4) & Mid(txtFechaInicioPresSaludEmergencia.Text, 4, 2)
            If oRsTabla1.RecordCount > 0 Then
                oRsTabla1.MoveFirst
                Do While Not oRsTabla1.EOF
                    lcLineaTxtPlano = ""
                    lcLineaTxtPlano = lcLineaTxtPlano & lcPeriodo & "|"
                    lcLineaTxtPlano = lcLineaTxtPlano & lcIpress & "|"
                    lcLineaTxtPlano = lcLineaTxtPlano & lcUgipress & "|"
                    lcLineaTxtPlano = lcLineaTxtPlano & oRsTabla1!Sexo & "|"
                    lcLineaTxtPlano = lcLineaTxtPlano & oRsTabla1!Grupo & "|"
                    lcLineaTxtPlano = lcLineaTxtPlano & Trim(str(oRsTabla1!Atenciones)) & "|"
                    lcLineaTxtPlano = lcLineaTxtPlano & Trim(str(oRsTabla1!Atendidos))
                    act.WriteLine (lcLineaTxtPlano)
                    
                    lnContadorDetalle = lnContadorDetalle + 1
                    DoEvents
                    progresSunasaDetalle.Value = lnContadorDetalle
                    Me.Refresh
                    
                    oRsTabla1.MoveNext
                Loop
            End If
            act.Close
            '
            lcLineaTxtPlano = lcIpress & "_" & Right(txtFechaInicioPresSaludEmergencia.Text, 4) & "_" & _
                              Mid(txtFechaInicioPresSaludEmergencia.Text, 4, 2) & "_TAC2.TXT"
            Set act = fso.CreateTextFile(lcParametro313 & lcLineaTxtPlano, True)
            progresSunasaDetalle.Min = 0
            progresSunasaDetalle.Max = oRsTabla2.RecordCount
            lnContadorDetalle = 0
            lcPeriodo = Right(txtFechaInicioPresSaludEmergencia.Text, 4) & Mid(txtFechaInicioPresSaludEmergencia.Text, 4, 2)
            If oRsTabla2.RecordCount > 0 Then
                oRsTabla2.MoveFirst
                Do While Not oRsTabla2.EOF
                    lcLineaTxtPlano = ""
                    lcLineaTxtPlano = lcLineaTxtPlano & lcPeriodo & "|"
                    lcLineaTxtPlano = lcLineaTxtPlano & lcIpress & "|"
                    lcLineaTxtPlano = lcLineaTxtPlano & lcUgipress & "|"
                    lcLineaTxtPlano = lcLineaTxtPlano & oRsTabla2!Sexo & "|"
                    lcLineaTxtPlano = lcLineaTxtPlano & oRsTabla2!Grupo & "|"
                    lcLineaTxtPlano = lcLineaTxtPlano & oRsTabla2!dx & "|"
                    lcLineaTxtPlano = lcLineaTxtPlano & Trim(str(oRsTabla2!Atendidos))
                    act.WriteLine (lcLineaTxtPlano)
                    
                    lnContadorDetalle = lnContadorDetalle + 1
                    DoEvents
                    progresSunasaDetalle.Value = lnContadorDetalle
                    Me.Refresh
                    
                    oRsTabla2.MoveNext
                Loop
            End If
            act.Close
        
        
    Else
            progresSunasaDetalle.Min = 0
            progresSunasaDetalle.Max = 1
            progresSunasaDetalle.Value = 1
            Me.Refresh
    End If
    rsReporte.Close
    Set rsReporte = Nothing
    Set oRsTmp2 = Nothing
    Set oRsTabla1 = Nothing
    Set oRsTabla2 = Nothing
    Set oRsTabla11 = Nothing
    Set oRsTabla22 = Nothing
    Set oRsFox = Nothing
End Sub



Sub GeneraTramaEmisionHospitalizacion2016(oConexion As ADODB.Connection, lcIpress As String, lcUgipress As String, _
                                              lcParametro313 As String)

Dim rsReporte As New Recordset
Dim oRsTmp2 As New Recordset
Dim oRsTabla1 As New Recordset
Dim oRsTabla11 As New Recordset
Dim oRsTabla2 As New Recordset
Dim oRsTabla22 As New Recordset
Dim oRsFox As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim lcSql As String
Dim lcLineaTxtPlano As String
Dim lcDxAtencion As String
Dim lnContadorDetalle As Long, lcPeriodo As String, lnIdAtencion As Long, lnNuevo As Boolean, lnIdTipoServicio As Integer
Dim lcSexo As String, lcGrupo As String, lcDx As String, lcColegio As String
Dim lbPersonaA As Boolean, lbPersonaDx As Boolean, lcServicio As String, lcErrores As String, lcHoraEstanciaMax As String
Dim lbFallecido As Boolean, lnCamasTotales As Integer, lnCamasOcupadasYvacantes As Integer, lnIdServicio As Long
Dim lnDiasEstancia As Integer, mda_FechaInicio As Date, mda_FechaFin As Date
        lblTabla.Caption = chbPresSaludhospitalizacion.Caption
        mda_FechaInicio = CDate(txtFechaInicioPresSaludHospitalizacion.Text)
        mda_FechaFin = CDate(txtFechaFinPresSaludHospitalizacion.Text)
        lnIdTipoServicio = sghTipoServicio.sghHospitalizacion
        lcErrores = ""
        lcHoraEstanciaMax = lcBuscaParametro.SeleccionaFilaParametro(201)
        Set rsReporte = mo_AdminReportes.ReporteIngresosHospitalarios(0, 0, 0, mda_FechaInicio, mda_FechaFin, _
                                                                                        lnIdTipoServicio)
        If rsReporte.RecordCount > 0 Then
           With oRsTabla1
                  .Fields.Append "servicio", adVarChar, 7
                  .Fields.Append "ingresos", adInteger
                  .Fields.Append "egresos", adInteger
                  .Fields.Append "estancias", adInteger
                  .Fields.Append "PacientesDias", adInteger
                  .Fields.Append "camas", adInteger
                  .Fields.Append "DiasCamas", adInteger
                  .Fields.Append "Fallecidos", adInteger
                  .Fields.Append "camasIdServicios", adVarChar, 500
                  .LockType = adLockOptimistic
                  .Open
           End With
           progresSunasaDetalle.Min = 0
           progresSunasaDetalle.Max = rsReporte.RecordCount
           lnContadorDetalle = 0
           
           rsReporte.MoveFirst
           Do While Not rsReporte.EOF
              lcServicio = IIf(IsNull(rsReporte!CodServSSegr), IIf(IsNull(rsReporte!CodServSSing), "", rsReporte!CodServSSing), rsReporte!CodServSSegr)
              lnIdServicio = IIf(IsNull(rsReporte!CodServSSegr), IIf(IsNull(rsReporte!CodServSSing), 0, rsReporte!IdServicioIngreso), rsReporte!IdServicioEgreso)
              If lcServicio <> "" Then
                    lnDiasEstancia = 0
                    If Not IsNull(rsReporte!FechaEgreso) Then
                       lnDiasEstancia = lcBuscaParametro.DiasDelPacienteEnHospitalizacionEmergencia(rsReporte!FechaIngreso, rsReporte!HoraIngreso, rsReporte!FechaEgreso, rsReporte!horaEgreso, lcHoraEstanciaMax)
                    End If
                    '
                    lbFallecido = False
                    If Not IsNull(rsReporte!CondicionAlta) Then
                    If rsReporte!CondicionAlta = 4 Then
                       lbFallecido = True
                    End If
                    End If
                    '
                    lnNuevo = True
                    If oRsTabla1.RecordCount > 0 Then
                       oRsTabla1.MoveFirst
                       oRsTabla1.Find "servicio='" & lcServicio & "'"
                       If Not oRsTabla1.EOF Then
                          lnNuevo = False
                       End If
                    End If
                    If lnNuevo = True Then
                       Set oRsTmp2 = mo_ReglasHoteleria.CamasSeleccionarPorIdServicio(lnIdServicio, oConexion)
                       lnCamasTotales = oRsTmp2.RecordCount
                       lnCamasOcupadasYvacantes = lnCamasTotales * sighentidades.DevuelveUltimoDiaDelMes(Month(mda_FechaInicio), Year(mda_FechaInicio))
                       oRsTmp2.Close
                       '
                       oRsTabla1.AddNew
                       oRsTabla1.Fields!Servicio = lcServicio
                       oRsTabla1.Fields!ingresos = 1
                       If Not IsNull(rsReporte!FechaEgreso) Then
                          oRsTabla1.Fields!egresos = 1
                          oRsTabla1.Fields!estancias = lnDiasEstancia
                       End If
                       oRsTabla1.Fields!pacientesDias = 1
                       oRsTabla1.Fields!camas = lnCamasTotales
                       oRsTabla1.Fields!diasCamas = lnCamasOcupadasYvacantes
                       If lbFallecido = True Then
                          oRsTabla1.Fields!facellidos = 1
                       End If
                       oRsTabla1!camasIdServicios = "/" & Trim(str(lnIdServicio)) & "/"
                    Else
                       If InStr(oRsTabla1!camasIdServicios, "/" & Trim(str(lnIdServicio)) & "/") = 0 Then
                            Set oRsTmp2 = mo_ReglasHoteleria.CamasSeleccionarPorIdServicio(lnIdServicio, oConexion)
                            lnCamasTotales = oRsTabla1.Fields!camas + oRsTmp2.RecordCount
                            lnCamasOcupadasYvacantes = lnCamasTotales * sighentidades.DevuelveUltimoDiaDelMes(Month(mda_FechaInicio), Year(mda_FechaInicio))
                            oRsTmp2.Close
                            oRsTabla1.Fields!camas = lnCamasTotales
                            oRsTabla1.Fields!diasCamas = lnCamasOcupadasYvacantes
                            oRsTabla1!camasIdServicios = oRsTabla1!camasIdServicios & Trim(str(lnIdServicio)) & "/"
                       End If
                       oRsTabla1.Fields!ingresos = oRsTabla1.Fields!ingresos + 1
                       If Not IsNull(rsReporte!FechaEgreso) Then
                          oRsTabla1.Fields!egresos = oRsTabla1.Fields!egresos + 1
                          oRsTabla1.Fields!estancias = oRsTabla1.Fields!estancias + lnDiasEstancia
                       End If
                       oRsTabla1.Fields!pacientesDias = oRsTabla1.Fields!pacientesDias + 1
                       If lbFallecido = True Then
                          oRsTabla1.Fields!facellidos = oRsTabla1.Fields!facellidos + 1
                       End If
                       
                    End If
                    
              Else
                    If IsNull(rsReporte!CodServSSing) Then
                      If InStr(lcErrores, rsReporte!servicioIngreso) = 0 Then
                         lcErrores = lcErrores & " Falta Configurar para " & rsReporte!servicioIngreso & " el CODIGO UPS SUSALUD en opción GENERAL->SERVICIOS" & Chr(13)
                      End If
                    ElseIf IsNull(rsReporte!CodServSSegr) Then
                      If InStr(lcErrores, rsReporte!ServicioEgreso) = 0 Then
                         lcErrores = lcErrores & " Falta Configurar para " & rsReporte!ServicioEgreso & " el CODIGO UPS SUSALUD en opción GENERAL->SERVICIOS" & Chr(13)
                      End If
                    End If
              End If
           
              DoEvents
              progresSunasaDetalle.Value = lnContadorDetalle
              Me.Refresh
              lnContadorDetalle = lnContadorDetalle + 1
           
              rsReporte.MoveNext
           Loop
        
        End If
        
        
        Set rsReporte = mo_AdminReportes.ReporteEgresosHospitalarios(0, 0, 0, mda_FechaInicio, mda_FechaFin, _
                                                                            lnIdTipoServicio)
        If rsReporte.RecordCount > 0 Then
           With oRsTabla2
                  .Fields.Append "sexo", adVarChar, 1
                  .Fields.Append "grupo", adVarChar, 2
                  .Fields.Append "dx", adVarChar, 5
                  .Fields.Append "atendidos", adInteger
                  .LockType = adLockOptimistic
                  .Open
           End With
           With oRsTabla22
                  .Fields.Append "sexo", adVarChar, 1
                  .Fields.Append "grupo", adVarChar, 2
                  .Fields.Append "dx", adVarChar, 5
                  .Fields.Append "NroHistoriaClinica", adInteger
                  .Fields.Append "TipoNumeracion", adVarChar, 100
                  .LockType = adLockOptimistic
                  .Open
           End With
           progresSunasaDetalle.Min = 0
           progresSunasaDetalle.Max = rsReporte.RecordCount
           lnContadorDetalle = 0
           
           rsReporte.MoveFirst
           Do While Not rsReporte.EOF
              lcSexo = IIf(UCase(Left(rsReporte!Sexo, 1)) = "M", "1", "2")
              lcGrupo = ""
              If rsReporte!TipoEdad = "A" Then
                 If rsReporte!Edad >= 1 And rsReporte!Edad <= 4 Then
                    lcGrupo = "2"
                 ElseIf rsReporte!Edad >= 5 And rsReporte!Edad <= 9 Then
                    lcGrupo = "3"
                 ElseIf rsReporte!Edad >= 10 And rsReporte!Edad <= 14 Then
                    lcGrupo = "4"
                 ElseIf rsReporte!Edad >= 15 And rsReporte!Edad <= 19 Then
                    lcGrupo = "5"
                 ElseIf rsReporte!Edad >= 20 And rsReporte!Edad <= 24 Then
                    lcGrupo = "6"
                 ElseIf rsReporte!Edad >= 25 And rsReporte!Edad <= 29 Then
                    lcGrupo = "7"
                 ElseIf rsReporte!Edad >= 30 And rsReporte!Edad <= 34 Then
                    lcGrupo = "8"
                 ElseIf rsReporte!Edad >= 35 And rsReporte!Edad <= 39 Then
                    lcGrupo = "9"
                 ElseIf rsReporte!Edad >= 40 And rsReporte!Edad <= 44 Then
                    lcGrupo = "10"
                 ElseIf rsReporte!Edad >= 45 And rsReporte!Edad <= 49 Then
                    lcGrupo = "11"
                 ElseIf rsReporte!Edad >= 50 And rsReporte!Edad <= 54 Then
                    lcGrupo = "12"
                 ElseIf rsReporte!Edad >= 55 And rsReporte!Edad <= 59 Then
                    lcGrupo = "13"
                 ElseIf rsReporte!Edad >= 60 And rsReporte!Edad <= 64 Then
                    lcGrupo = "14"
                 Else
                    lcGrupo = "15"
                 End If
              Else
                 lcGrupo = "1"
              End If
              '
              lcDx = ""
              If lnIdTipoServicio <> sghTipoServicio.sghConsultaExterna Then
                    Set oRsTmp2 = mo_AdminReportes.ReporteAtencionesDiagnosticosDeEgreso(rsReporte!idAtencion)
              Else
                    Set oRsTmp2 = mo_ReglasAdmision.BuscaCEAtencionesDx(rsReporte!idAtencion)
              End If
              oRsTmp2.Filter = "codigoDx<>null"
              If oRsTmp2.RecordCount > 0 Then
                   lcDx = Left(oRsTmp2!CodigoDx, 5)
              End If
              oRsTmp2.Close

              'Personas dx
              If lcDx <> "" Then
                    lbPersonaDx = True
                    If oRsTabla22.RecordCount > 0 Then
                       oRsTabla22.MoveFirst
                       Do While Not oRsTabla22.EOF
                          If oRsTabla22!Sexo = lcSexo And oRsTabla22!Grupo = lcGrupo And oRsTabla22!dx = lcDx And _
                                                          oRsTabla22!NroHistoriaClinica = rsReporte!NroHistoriaClinica And _
                                                          oRsTabla22!TipoNumeracion = rsReporte!TipoNumeracion Then
                             lbPersonaDx = False
                             Exit Do
                          End If
                          oRsTabla22.MoveNext
                       Loop
                    End If
                    If lbPersonaDx = True Then
                       oRsTabla22.AddNew
                       oRsTabla22.Fields!Sexo = lcSexo
                       oRsTabla22.Fields!Grupo = lcGrupo
                       oRsTabla22.Fields!dx = lcDx
                       oRsTabla22.Fields!NroHistoriaClinica = rsReporte!NroHistoriaClinica
                       oRsTabla22.Fields!TipoNumeracion = rsReporte!TipoNumeracion
                       oRsTabla22.Update
                    End If
                    'dx
                    lnNuevo = True
                    If oRsTabla2.RecordCount > 0 Then
                       oRsTabla2.MoveFirst
                       Do While Not oRsTabla2.EOF
                          If oRsTabla2!Sexo = lcSexo And oRsTabla2!Grupo = lcGrupo And oRsTabla2.Fields!dx = lcDx Then
                             lnNuevo = False
                             Exit Do
                          End If
                          oRsTabla2.MoveNext
                       Loop
                    End If
                    If lnNuevo = True Then
                          oRsTabla2.AddNew
                          oRsTabla2.Fields!Sexo = lcSexo
                          oRsTabla2.Fields!Grupo = lcGrupo
                          oRsTabla2.Fields!dx = lcDx
                          If lbPersonaDx = True Then
                             oRsTabla2.Fields!Atendidos = 1
                          End If
                    Else
                          If lbPersonaDx = True Then
                             oRsTabla2.Fields!Atendidos = oRsTabla2.Fields!Atendidos + 1
                          End If
                    End If
                    oRsTabla2.Update
              End If
              '
              DoEvents
              progresSunasaDetalle.Value = lnContadorDetalle
              Me.Refresh
              lnContadorDetalle = lnContadorDetalle + 1
            
              rsReporte.MoveNext
           Loop
    Else
            progresSunasaDetalle.Min = 0
            progresSunasaDetalle.Max = 1
            progresSunasaDetalle.Value = 1
            Me.Refresh
    End If
    rsReporte.Close
    
    If lcErrores <> "" Then
       MsgBox lcErrores, vbInformation, Me.Caption
    End If
    
    Dim fso
    Dim act
    Set fso = CreateObject("scripting.filesystemobject")
    '
    lcLineaTxtPlano = lcIpress & "_" & Right(txtFechaInicioPresSaludHospitalizacion.Text, 4) & "_" & _
                      Mid(txtFechaInicioPresSaludHospitalizacion.Text, 4, 2) & "_TAD1.TXT"
    Set act = fso.CreateTextFile(lcParametro313 & lcLineaTxtPlano, True)
    progresSunasaDetalle.Min = 0
    progresSunasaDetalle.Max = IIf(oRsTabla1.RecordCount = 0, 2, oRsTabla1.RecordCount)
    lnContadorDetalle = 0
    lcPeriodo = Right(txtFechaInicioPresSaludHospitalizacion.Text, 4) & Mid(txtFechaInicioPresSaludHospitalizacion.Text, 4, 2)
    If oRsTabla1.RecordCount > 0 Then
        oRsTabla1.MoveFirst
        Do While Not oRsTabla1.EOF
            lcLineaTxtPlano = ""
            lcLineaTxtPlano = lcLineaTxtPlano & lcPeriodo & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & lcIpress & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & lcUgipress & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & oRsTabla1!Servicio & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & oRsTabla1!ingresos & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & Trim(str(oRsTabla1!egresos)) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & Trim(str(oRsTabla1!estancias)) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & Trim(str(oRsTabla1!pacientesDias)) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & Trim(str(oRsTabla1!camas)) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & Trim(str(oRsTabla1!diasCamas)) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & Trim(str(oRsTabla1!fallecidos))
            act.WriteLine (lcLineaTxtPlano)
            
            lnContadorDetalle = lnContadorDetalle + 1
            DoEvents
            progresSunasaDetalle.Value = lnContadorDetalle
            Me.Refresh
            
            oRsTabla1.MoveNext
        Loop
    End If
    act.Close
    '
    lcLineaTxtPlano = lcIpress & "_" & Right(txtFechaInicioPresSaludHospitalizacion.Text, 4) & "_" & _
                      Mid(txtFechaInicioPresSaludHospitalizacion.Text, 4, 2) & "_TAD2.TXT"
    Set act = fso.CreateTextFile(lcParametro313 & lcLineaTxtPlano, True)
    progresSunasaDetalle.Min = 0
    progresSunasaDetalle.Max = oRsTabla2.RecordCount
    lnContadorDetalle = 0
    lcPeriodo = Right(txtFechaInicioPresSaludHospitalizacion.Text, 4) & Mid(txtFechaInicioPresSaludHospitalizacion.Text, 4, 2)
    If oRsTabla2.RecordCount > 0 Then
        oRsTabla2.MoveFirst
        Do While Not oRsTabla2.EOF
            lcLineaTxtPlano = ""
            lcLineaTxtPlano = lcLineaTxtPlano & lcPeriodo & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & lcIpress & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & lcUgipress & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & oRsTabla2!Sexo & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & oRsTabla2!Grupo & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & oRsTabla2!dx & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & Trim(str(oRsTabla2!Atendidos))
            act.WriteLine (lcLineaTxtPlano)
            
            lnContadorDetalle = lnContadorDetalle + 1
            DoEvents
            progresSunasaDetalle.Value = lnContadorDetalle
            Me.Refresh
            
            oRsTabla2.MoveNext
        Loop
    End If
    act.Close
    
    
    Set rsReporte = Nothing
    Set oRsTmp2 = Nothing
    Set oRsTabla1 = Nothing
    Set oRsTabla2 = Nothing
    Set oRsTabla11 = Nothing
    Set oRsTabla22 = Nothing
    Set oRsFox = Nothing
End Sub





Sub GeneraTramaPrestacionesSaludCPT2016(oConexion As ADODB.Connection, lcIpress As String, lcUgipress As String, _
                                              lcParametro313 As String)
Dim oRsTmp As New Recordset
Dim oRsTmp2 As New Recordset
Dim oRsTabla As New Recordset
Dim oRsFox As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim lcSql As String
Dim lcLineaTxtPlano As String
Dim lcDxAtencion As String
Dim lnContadorDetalle As Long, lcPeriodo As String, lnIdAtencion As Long
Dim lnVivos As Long, lnMuertos As Long, lcTipoParto As String, lcComplicacion As String, lnNuevo As Boolean
Dim lnDiasEstancia As Integer, mda_FechaInicio As Date, mda_FechaFin As Date, lcErrores As String
        
    mda_FechaInicio = CDate(txtFcptInicio.Text)
    mda_FechaFin = CDate(txtFcptFin.Text)
    lblTabla.Caption = chbProcedimientos.Caption
    
    Set oRsTmp = mo_ReglasFacturacion.FacturacionServicioDespachoXfechas(mda_FechaInicio, mda_FechaFin, oConexion)
    If oRsTmp.RecordCount > 0 Then
        With oRsTabla
              .Fields.Append "cpt", adVarChar, 10, adFldIsNullable
              .Fields.Append "cptNumero", adInteger
              .Fields.Append "servicio", adVarChar, 7, adFldIsNullable
              .LockType = adLockOptimistic
              .Open
        End With
        progresSunasaDetalle.Min = 0
        progresSunasaDetalle.Max = oRsTmp.RecordCount
        lnContadorDetalle = 0
        lcErrores = ""
        oRsTmp.MoveFirst
        Do While Not oRsTmp.EOF
           If IsNull(oRsTmp!codigoServicioSuSalud) Then
              If InStr(lcErrores, oRsTmp!DServicio) = 0 Then
                If oRsTmp!DServicio <> "" Then
                   lcErrores = lcErrores & " Falta Configurar para " & _
                              IIf(oRsTmp!idTipoServicio = 1, "(CE) ", IIf(oRsTmp!idTipoServicio = 2, "(Emer) ", "(Hosp) ")) & _
                              oRsTmp!DServicio & " el CODIGO UPS SUSALUD en opción GENERAL->SERVICIOS" & Chr(13)
                End If
              End If
           Else
                lnNuevo = True
                If oRsTabla.RecordCount > 0 Then
                   oRsTabla.MoveFirst
                   Do While Not oRsTabla.EOF
                      If oRsTabla!cpt = oRsTmp!cpt And oRsTabla!Servicio = oRsTmp!codigoServicioSuSalud Then
                         lnNuevo = False
                         Exit Do
                      End If
                      oRsTabla.MoveNext
                   Loop
                End If
                If lnNuevo = True Then
                   oRsTabla.AddNew
                   oRsTabla.Fields!cpt = oRsTmp!cpt
                   oRsTabla.Fields!cptNumero = oRsTmp!Cantidad
                   oRsTabla.Fields!Servicio = oRsTmp!codigoServicioSuSalud
                Else
                   oRsTabla.Fields!cptNumero = oRsTabla!cptNumero + oRsTmp!Cantidad
                End If
                oRsTabla.Update
           End If
           
           DoEvents
           progresSunasaDetalle.Value = lnContadorDetalle
           Me.Refresh
           lnContadorDetalle = lnContadorDetalle + 1
           
           oRsTmp.MoveNext
        Loop
        If lcErrores <> "" Then
            MsgBox lcErrores, vbInformation, Me.Caption
        End If
        Dim fso
        Dim act
        Set fso = CreateObject("scripting.filesystemobject")
        lcLineaTxtPlano = lcIpress & "_" & Right(txtFcptInicio.Text, 4) & "_" & _
                          Mid(txtFcptInicio.Text, 4, 2) & "_TAG0.TXT"
        Set act = fso.CreateTextFile(lcParametro313 & lcLineaTxtPlano, True)
        progresSunasaDetalle.Min = 0
        progresSunasaDetalle.Max = oRsTabla.RecordCount
        lnContadorDetalle = 0
        lcPeriodo = Right(txtFcptInicio.Text, 4) & Mid(txtFcptInicio.Text, 4, 2)
        oRsTabla.MoveFirst
        Do While Not oRsTabla.EOF
            lcLineaTxtPlano = ""
            lcLineaTxtPlano = lcLineaTxtPlano & lcPeriodo & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & lcIpress & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & lcUgipress & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & oRsTabla!cpt & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & Trim(str(oRsTabla!cptNumero)) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & oRsTabla!Servicio
            act.WriteLine (lcLineaTxtPlano)
            
            lnContadorDetalle = lnContadorDetalle + 1
            DoEvents
            progresSunasaDetalle.Value = lnContadorDetalle
            Me.Refresh
            
            oRsTabla.MoveNext
        Loop
        act.Close
    Else
        progresSunasaDetalle.Min = 0
        progresSunasaDetalle.Max = 1
        progresSunasaDetalle.Value = 1
        Me.Refresh
   End If
   Set oRsTmp = Nothing
   Set oRsTmp2 = Nothing
   Set oRsTabla = Nothing
   Set oRsFox = Nothing
End Sub


Sub GeneraTramaReferencias2016(oConexion As ADODB.Connection, lcIpress As String, lcUgipress As String, _
                                              lcParametro313 As String)
Dim oRsTmp As New Recordset
Dim oRsTmp2 As New Recordset
Dim oRsTabla As New Recordset
Dim oRsFox As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim lcSql As String
Dim lcLineaTxtPlano As String
Dim lcDxAtencion As String
Dim lnContadorDetalle As Long, lcPeriodo As String, lnIdAtencion As Long
Dim lnVivos As Long, lnMuertos As Long, lcTipoParto As String, lcComplicacion As String, lnNuevo As Boolean
Dim lnDiasEstancia As Integer, mda_FechaInicio As Date, mda_FechaFin As Date, lcErrores As String
Dim lcDxPrincipal As String, lcDxPrincipalTipo As String, lcDxSecundario As String, lcDxSecundarioTipo As String
    sighentidades.ParaAuditoria = "inicio"
    mda_FechaInicio = CDate(txtFreferInicio.Text)
    mda_FechaFin = CDate(txtFreferFin.Text)
    lblTabla.Caption = chbReferencias.Caption
    
    Set oRsTmp = mo_ReglasAdmision.AtencionesDatosAdicionalesXfechas(mda_FechaInicio, mda_FechaFin, oConexion)
    If oRsTmp.RecordCount > 0 Then
        With oRsTabla
              .Fields.Append "eessOrigenRenaes", adVarChar, 10, adFldIsNullable
              .Fields.Append "nroHistoriaClinica", adInteger
              .Fields.Append "dniTipo", adVarChar, 1, adFldIsNullable
              .Fields.Append "dniNumero", adVarChar, 12, adFldIsNullable
              .Fields.Append "sexo", adVarChar, 1, adFldIsNullable
              .Fields.Append "edad", adVarChar, 5, adFldIsNullable
              .Fields.Append "eessOrigenServicio", adVarChar, 7, adFldIsNullable
              .Fields.Append "eessDestinoRenaes", adVarChar, 10, adFldIsNullable
              .Fields.Append "eessDestinoServicio", adVarChar, 7, adFldIsNullable
              .Fields.Append "DxPrincipal", adVarChar, 5, adFldIsNullable
              .Fields.Append "DxPrincipalTipo", adVarChar, 2, adFldIsNullable
              .Fields.Append "DxSecundario", adVarChar, 5, adFldIsNullable
              .Fields.Append "DxSecundarioTipo", adVarChar, 2, adFldIsNullable
              .Fields.Append "fExtension", adVarChar, 8, adFldIsNullable
              .Fields.Append "fTramite", adVarChar, 8, adFldIsNullable
              .LockType = adLockOptimistic
              .Open
        End With
        sighentidades.ParaAuditoria = "genera temporal"
        progresSunasaDetalle.Min = 0
        progresSunasaDetalle.Max = oRsTmp.RecordCount
        sighentidades.ParaAuditoria = "progressbar"
        lnContadorDetalle = 0
        lcErrores = ""
        oRsTmp.MoveFirst
        Do While Not oRsTmp.EOF
        
            lcDxPrincipal = "": lcDxPrincipalTipo = "": lcDxSecundario = "": lcDxSecundarioTipo = ""
            If oRsTmp!idTipoServicio <> sghTipoServicio.sghConsultaExterna Then
                Set oRsTmp2 = mo_AdminReportes.ReporteAtencionesDiagnosticosDeEgreso(oRsTmp!idAtencion)
            Else
                Set oRsTmp2 = mo_ReglasAdmision.BuscaCEAtencionesDx(oRsTmp!idAtencion)
            End If
            If oRsTmp2.RecordCount > 0 Then
               oRsTmp2.MoveFirst
               Do While Not oRsTmp2.EOF
                  If Not IsNull(oRsTmp2!CodigoDx) Then
                        If oRsTmp2!IdSubclasificacionDx = 102 Or oRsTmp2!IdSubclasificacionDx = 301 Or _
                                       oRsTmp2!IdSubclasificacionDx = 303 Or oRsTmp2!IdSubclasificacionDx = 402 Then
                           lcDxPrincipal = Left(oRsTmp2!CodigoDx, 5)
                           lcDxPrincipalTipo = "02"
                        Else
                           lcDxSecundario = Left(oRsTmp2!CodigoDx, 5)
                           lcDxSecundarioTipo = "01"
                        End If
                  End If
                  oRsTmp2.MoveNext
               Loop
            End If
            oRsTmp2.Close
            sighentidades.ParaAuditoria = "dx"
            oRsTabla.AddNew
            sighentidades.ParaAuditoria = "1 " & Trim(str(oRsTmp!idAtencion))
            oRsTabla.Fields!NroHistoriaClinica = Val(HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(str(oRsTmp!NroHistoriaClinica)), False))
            sighentidades.ParaAuditoria = "2"
            If oRsTmp!IdDocIdentidad = 1 Or oRsTmp!IdDocIdentidad = 2 Or oRsTmp!IdDocIdentidad = 3 Or oRsTmp!IdDocIdentidad = 5 Then
                oRsTabla.Fields!dniTipo = IIf(IsNull(oRsTmp!IdDocIdentidad), "", Trim(str(oRsTmp!IdDocIdentidad)))
                oRsTabla.Fields!dniNumero = IIf(IsNull(oRsTmp!nroDocumento), "", Left(str(oRsTmp!nroDocumento), 12))
                sighentidades.ParaAuditoria = "3"
            Else
                oRsTabla.Fields!dniTipo = ""
                oRsTabla.Fields!dniNumero = ""
                sighentidades.ParaAuditoria = "4"
            End If
            oRsTabla.Fields!eessOrigenRenaes = IIf(IsNull(oRsTmp!eessOrigenRenaes), "", oRsTmp!eessOrigenRenaes)
            sighentidades.ParaAuditoria = "5"
            oRsTabla.Fields!Sexo = Trim(str(oRsTmp!idTipoSexo))
            sighentidades.ParaAuditoria = "6"
            oRsTabla.Fields!Edad = Trim(str(oRsTmp!Edad)) & "-" & sighentidades.EdadDevuelveTipo(oRsTmp!idTipoEdad)
            sighentidades.ParaAuditoria = "7"
            oRsTabla.Fields!eessOrigenServicio = IIf(IsNull(oRsTmp!referenciaOservicio), "", oRsTmp!referenciaOservicio)
            sighentidades.ParaAuditoria = "8"
            oRsTabla.Fields!EESSdestinoRenaes = IIf(IsNull(oRsTmp!EESSdestinoRenaes), "", oRsTmp!EESSdestinoRenaes)
            sighentidades.ParaAuditoria = "9"
            oRsTabla.Fields!eessDestinoServicio = IIf(IsNull(oRsTmp!referenciaDservicio), "", oRsTmp!referenciaDservicio)
            sighentidades.ParaAuditoria = "10"
            oRsTabla.Fields!DxPrincipal = lcDxPrincipal
            sighentidades.ParaAuditoria = "11"
            oRsTabla.Fields!DxPrincipalTipo = lcDxPrincipalTipo
            sighentidades.ParaAuditoria = "12"
            oRsTabla.Fields!DxSecundario = lcDxSecundario
            sighentidades.ParaAuditoria = "13"
            oRsTabla.Fields!DxSecundarioTipo = lcDxSecundarioTipo
            sighentidades.ParaAuditoria = "14"
            oRsTabla.Fields!fExtension = IIf(IsNull(oRsTmp!referenciaDfextension), "", Format(oRsTmp!referenciaDfextension, sighentidades.DevuelveFechaSoloFormato_AAAAMMDD))
            sighentidades.ParaAuditoria = "15"
            oRsTabla.Fields!fTramite = IIf(IsNull(oRsTmp!referenciaDftramite), "", Format(oRsTmp!referenciaDftramite, sighentidades.DevuelveFechaSoloFormato_AAAAMMDD))
            sighentidades.ParaAuditoria = "16"
            oRsTabla.Update
            sighentidades.ParaAuditoria = "agrega 1item a la tabla"
            DoEvents
            progresSunasaDetalle.Value = lnContadorDetalle
            Me.Refresh
            lnContadorDetalle = lnContadorDetalle + 1
           
            oRsTmp.MoveNext
            
        Loop
        sighentidades.ParaAuditoria = "fso"
        Dim fso
        Dim act
        Set fso = CreateObject("scripting.filesystemobject")
        sighentidades.ParaAuditoria = "crea objeto"
        lcLineaTxtPlano = lcIpress & "_" & Right(txtFreferInicio.Text, 4) & "_" & _
                          Mid(txtFreferInicio.Text, 4, 2) & "_TAI0.TXT"
sighentidades.ParaAuditoria = "line texto"
        Set act = fso.CreateTextFile(lcParametro313 & lcLineaTxtPlano, True)
sighentidades.ParaAuditoria = "crea texto"
        progresSunasaDetalle.Min = 0
        progresSunasaDetalle.Max = oRsTabla.RecordCount
        lnContadorDetalle = 1
        lcPeriodo = Right(txtFreferInicio.Text, 4) & Mid(txtFreferInicio.Text, 4, 2)
sighentidades.ParaAuditoria = "periodo"
        oRsTabla.MoveFirst
        Do While Not oRsTabla.EOF
            
            lcLineaTxtPlano = ""
            lcLineaTxtPlano = lcLineaTxtPlano & lcPeriodo & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & lcIpress & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & lcUgipress & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & oRsTabla!eessOrigenRenaes & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & Right("00000" & Trim(str(lnContadorDetalle)), 5) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & Trim(str(oRsTabla!NroHistoriaClinica)) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & oRsTabla!dniTipo & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & Trim(oRsTabla!dniNumero) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & oRsTabla!Sexo & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & oRsTabla!Edad & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & oRsTabla!eessOrigenServicio & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & oRsTabla!EESSdestinoRenaes & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & oRsTabla!eessDestinoServicio & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & oRsTabla!DxPrincipal & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & oRsTabla!DxPrincipalTipo & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & oRsTabla!DxSecundario & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & oRsTabla!DxSecundarioTipo & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & oRsTabla!fExtension & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & oRsTabla!fTramite
            act.WriteLine (lcLineaTxtPlano)
            
            sighentidades.ParaAuditoria = "graba linea"
            lnContadorDetalle = lnContadorDetalle + 1
            DoEvents
            progresSunasaDetalle.Value = lnContadorDetalle - 1
            Me.Refresh
            sighentidades.ParaAuditoria = "progresbar2"
            oRsTabla.MoveNext
        Loop
        act.Close
        sighentidades.ParaAuditoria = "cierra texto"
    Else
        sighentidades.ParaAuditoria = "antes progresbar2 else"
        progresSunasaDetalle.Min = 0
        progresSunasaDetalle.Max = 1
        progresSunasaDetalle.Value = 1
        Me.Refresh
        sighentidades.ParaAuditoria = "paso progresbar2 else"
   End If
   sighentidades.ParaAuditoria = "antes nothing"
   Set oRsTmp = Nothing
   Set oRsTmp2 = Nothing
   Set oRsTabla = Nothing
   Set oRsFox = Nothing
   sighentidades.ParaAuditoria = "okey"
End Sub




Sub GeneraTramaProgramacion2016(oConexion As ADODB.Connection, lcIpress As String, lcUgipress As String, _
                                              lcParametro313 As String)
Dim oRsTmp As New Recordset
Dim oRsTmp2 As New Recordset
Dim oRsTabla As New Recordset
Dim oRsFox As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim lcSql As String
Dim lcLineaTxtPlano As String
Dim lcDxAtencion As String
Dim lnContadorDetalle As Long, lcPeriodo As String, lnIdAtencion As Long
Dim lnVivos As Long, lnMuertos As Long, lcTipoParto As String, lcComplicacion As String, lnNuevo As Boolean
Dim lnDiasEstancia As Integer, mda_FechaInicio As Date, mda_FechaFin As Date, lcErrores As String
Dim lcDxPrincipal As String, lcDxPrincipalTipo As String, lcDxSecundario As String, lcDxSecundarioTipo As String
Dim ldFechaHrInicial As Date, ldFechaHrFinal As Date, lnHoras As Integer, lnIdEspecCObst As Long, lnIdEspecCQuirur As Long
Dim lnTotalProfesionales As Long
    mda_FechaInicio = CDate(txtFProgInicio.Text)
    mda_FechaFin = CDate(txtFprogFinal.Text)
    lblTabla.Caption = chbProgAsistencial.Caption
    lnIdEspecCObst = Val(lcBuscaParametro.SeleccionaFilaParametro(503))
    lnIdEspecCQuirur = Val(lcBuscaParametro.SeleccionaFilaParametro(504))
    Set oRsTmp = mo_ReglasDeProgMedica.ProgramacionMedicaSeleccionarXfechas(mda_FechaInicio, mda_FechaFin, oConexion)
    If oRsTmp.RecordCount > 0 Then
        With oRsTabla
              .Fields.Append "codigoProfesional", adVarChar, 2, adFldIsNullable
              .Fields.Append "totalProfesionales", adInteger
              .Fields.Append "horasCE", adInteger
              .Fields.Append "horasEmergencia", adInteger
              .Fields.Append "horasHospitalizacion", adInteger
              .Fields.Append "horasAdministrativas", adInteger
              .Fields.Append "horasCapacitacion", adInteger
              .Fields.Append "horasCentroQx", adInteger
              .Fields.Append "horasCentroObst", adInteger
              .Fields.Append "horasProcedimientos", adInteger
              .Fields.Append "horasComplementarias", adInteger
              .LockType = adLockOptimistic
              .Open
        End With
        progresSunasaDetalle.Min = 0
        progresSunasaDetalle.Max = oRsTmp.RecordCount
        lnContadorDetalle = 0
        lcErrores = ""
        oRsTmp.MoveFirst
        Do While Not oRsTmp.EOF
            ldFechaHrInicial = CDate(Format(oRsTmp!fecha, sighentidades.DevuelveFechaSoloFormato_DMY) & " " & oRsTmp!HoraInicio)
            ldFechaHrFinal = CDate(Format(oRsTmp!fecha, sighentidades.DevuelveFechaSoloFormato_DMY) & " " & oRsTmp!HoraFin)
            lnHoras = DateDiff("h", ldFechaHrInicial, ldFechaHrFinal)
            lnNuevo = True
            If oRsTabla.RecordCount > 0 Then
               oRsTabla.MoveFirst
               Do While Not oRsTabla.EOF
                  If oRsTabla!codigoProfesional = oRsTmp!idColegioHIS Then
                     lnNuevo = False
                     Exit Do
                  End If
                  oRsTabla.MoveNext
               Loop
            End If
            If lnNuevo = True Then
                Set oRsTmp2 = mo_ReglasDeProgMedica.MedicosFiltrarDatos(" and medicos.idColegioHIS='" & oRsTmp!idColegioHIS & "'", oConexion)
                oRsTmp2.Filter = "nombre<>null"
                lnTotalProfesionales = oRsTmp2.RecordCount
                oRsTmp2.Close
                '
                oRsTabla.AddNew
                oRsTabla.Fields!codigoProfesional = oRsTmp!idColegioHIS
                oRsTabla.Fields!totalProfesionales = lnTotalProfesionales
                oRsTabla.Fields!horasCE = IIf(oRsTmp!idTipoServicio = 1, lnHoras, 0)
                oRsTabla.Fields!horasEmergencia = IIf(oRsTmp!idTipoServicio = 2, lnHoras, 0)
                oRsTabla.Fields!horasHospitalizacion = IIf(oRsTmp!idTipoServicio = 3, lnHoras, 0)
                oRsTabla.Fields!horasAdministrativas = IIf(oRsTmp!idTipoActividades = 4, lnHoras, 0)
                oRsTabla.Fields!horasCapacitacion = IIf(oRsTmp!idTipoActividades = 1, lnHoras, 0)
                oRsTabla.Fields!horasCentroQx = IIf(oRsTmp!IdEspecialidad = lnIdEspecCQuirur, lnHoras, 0)
                oRsTabla.Fields!horasCentroObst = IIf(oRsTmp!IdEspecialidad = lnIdEspecCObst, lnHoras, 0)
                oRsTabla.Fields!horasProcedimientos = IIf(oRsTmp!idTipoActividades = 2, lnHoras, 0)
                oRsTabla.Fields!horasComplementarias = IIf(oRsTmp!idTipoActividades = 3, lnHoras, 0)
            Else
                oRsTabla.Fields!horasCE = oRsTabla.Fields!horasCE + IIf(oRsTmp!idTipoServicio = 1, lnHoras, 0)
                oRsTabla.Fields!horasEmergencia = oRsTabla.Fields!horasEmergencia + IIf(oRsTmp!idTipoServicio = 2, lnHoras, 0)
                oRsTabla.Fields!horasHospitalizacion = oRsTabla.Fields!horasHospitalizacion + IIf(oRsTmp!idTipoServicio = 3, lnHoras, 0)
                oRsTabla.Fields!horasAdministrativas = oRsTabla.Fields!horasAdministrativas + IIf(oRsTmp!idTipoActividades = 4, lnHoras, 0)
                oRsTabla.Fields!horasCapacitacion = oRsTabla.Fields!horasCapacitacion + IIf(oRsTmp!idTipoActividades = 1, lnHoras, 0)
                oRsTabla.Fields!horasCentroQx = oRsTabla.Fields!horasCentroQx + IIf(oRsTmp!IdEspecialidad = lnIdEspecCQuirur, lnHoras, 0)
                oRsTabla.Fields!horasCentroObst = oRsTabla.Fields!horasCentroObst + IIf(oRsTmp!IdEspecialidad = lnIdEspecCObst, lnHoras, 0)
                oRsTabla.Fields!horasProcedimientos = oRsTabla.Fields!horasProcedimientos + IIf(oRsTmp!idTipoActividades = 2, lnHoras, 0)
                oRsTabla.Fields!horasComplementarias = oRsTabla.Fields!horasComplementarias + IIf(oRsTmp!idTipoActividades = 3, lnHoras, 0)
            End If
            oRsTabla.Update
           
            DoEvents
            progresSunasaDetalle.Value = lnContadorDetalle
            Me.Refresh
            lnContadorDetalle = lnContadorDetalle + 1
           
            oRsTmp.MoveNext
            
        Loop
        Dim fso
        Dim act
        Set fso = CreateObject("scripting.filesystemobject")
        lcLineaTxtPlano = lcIpress & "_" & Right(txtFProgInicio.Text, 4) & "_" & _
                          Mid(txtFProgInicio.Text, 4, 2) & "_TAJ0.TXT"
        Set act = fso.CreateTextFile(lcParametro313 & lcLineaTxtPlano, True)
        progresSunasaDetalle.Min = 0
        progresSunasaDetalle.Max = oRsTabla.RecordCount
        lnContadorDetalle = 0
        lcPeriodo = Right(txtFProgInicio.Text, 4) & Mid(txtFProgInicio.Text, 4, 2)
        oRsTabla.MoveFirst
        Do While Not oRsTabla.EOF
            lcLineaTxtPlano = ""
            lcLineaTxtPlano = lcLineaTxtPlano & lcPeriodo & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & lcIpress & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & lcUgipress & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & oRsTabla!codigoProfesional & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & Trim(str(oRsTabla!totalProfesionales)) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & Trim(str(oRsTabla!horasCE)) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & Trim(str(oRsTabla!horasEmergencia)) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & Trim(str(oRsTabla!horasHospitalizacion)) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & Trim(str(oRsTabla!horasAdministrativas)) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & Trim(str(oRsTabla!horasCapacitacion)) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & Trim(str(oRsTabla!horasCentroQx)) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & Trim(str(oRsTabla!horasCentroObst)) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & Trim(str(oRsTabla!horasProcedimientos)) & "|"
            lcLineaTxtPlano = lcLineaTxtPlano & Trim(str(oRsTabla!horasComplementarias))
            act.WriteLine (lcLineaTxtPlano)
            
            lnContadorDetalle = lnContadorDetalle + 1
            DoEvents
            progresSunasaDetalle.Value = lnContadorDetalle
            Me.Refresh
            
            oRsTabla.MoveNext
        Loop
        act.Close
    Else
        progresSunasaDetalle.Min = 0
        progresSunasaDetalle.Max = 1
        progresSunasaDetalle.Value = 1
        Me.Refresh
   End If
   Set oRsTmp = Nothing
   Set oRsTmp2 = Nothing
   Set oRsTabla = Nothing
   Set oRsFox = Nothing
End Sub








Private Sub progresSunasa1_Click()

End Sub
