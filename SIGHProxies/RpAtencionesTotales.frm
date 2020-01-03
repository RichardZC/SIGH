VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form RpAtencionesTotales 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Atenciones SIS "
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14550
   Icon            =   "RpAtencionesTotales.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   14550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   30
      TabIndex        =   28
      Top             =   7680
      Width           =   14475
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "RpAtencionesTotales.frx":0CCA
         DownPicture     =   "RpAtencionesTotales.frx":118E
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
         Left            =   7358
         Picture         =   "RpAtencionesTotales.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "RpAtencionesTotales.frx":1B66
         DownPicture     =   "RpAtencionesTotales.frx":1FC6
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
         Left            =   5828
         Picture         =   "RpAtencionesTotales.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame fraDatosHistoria 
      Caption         =   "Atenciones ANUALES"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7635
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   14505
      Begin VB.Frame Frame 
         Height          =   990
         Index           =   2
         Left            =   360
         TabIndex        =   58
         Top             =   6555
         Width           =   13920
         Begin Threed.SSOption optDemoraAtencion 
            Height          =   285
            Left            =   180
            TabIndex        =   59
            Top             =   585
            Width           =   6150
            _ExtentX        =   10848
            _ExtentY        =   503
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
            Caption         =   "Atenciones CE, tiempo de demora desde su CITA hasta atenciòn"
            Value           =   -1
         End
         Begin MSMask.MaskEdBox txtfCitaIni 
            Height          =   315
            Left            =   1680
            TabIndex        =   61
            Top             =   210
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
         Begin MSMask.MaskEdBox txtFcitaFin 
            Height          =   315
            Left            =   3105
            TabIndex        =   62
            Top             =   210
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
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Cita"
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
            Left            =   195
            TabIndex        =   63
            Top             =   225
            Width           =   840
         End
      End
      Begin VB.Frame Frame 
         Height          =   1365
         Index           =   1
         Left            =   345
         TabIndex        =   49
         Top             =   4905
         Width           =   13995
         Begin MSMask.MaskEdBox txtEdad 
            Height          =   330
            Left            =   11070
            TabIndex        =   57
            Top             =   900
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   2
            Mask            =   "##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFechaInicio 
            Height          =   315
            Left            =   1650
            TabIndex        =   50
            Top             =   225
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
         Begin MSMask.MaskEdBox txtFechaFin 
            Height          =   315
            Left            =   3075
            TabIndex        =   51
            Top             =   225
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
         Begin Threed.SSOption optMuerteEmergencia 
            Height          =   285
            Left            =   135
            TabIndex        =   53
            Top             =   600
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   503
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
            Caption         =   "Fallecidos en Emergencia"
            Value           =   -1
         End
         Begin Threed.SSOption optMuertesHosp 
            Height          =   285
            Left            =   4785
            TabIndex        =   54
            Top             =   600
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   503
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
            Caption         =   "Fallecidos en Hospitalización"
         End
         Begin Threed.SSOption optMuerteHospME 
            Height          =   285
            Left            =   9330
            TabIndex        =   55
            Top             =   585
            Width           =   4425
            _ExtentX        =   7805
            _ExtentY        =   503
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
            Caption         =   "Fallecidos en Hospitalización menores a una edad  "
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Edad hasta(años)"
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
            Left            =   9600
            TabIndex        =   56
            Top             =   930
            Width           =   1425
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Egreso"
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
            TabIndex        =   52
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         Height          =   3855
         Left            =   420
         TabIndex        =   1
         Top             =   630
         Width           =   14040
         Begin VB.Frame Frame1 
            Height          =   585
            Left            =   570
            TabIndex        =   14
            Top             =   1620
            Width           =   5310
            Begin Threed.SSOption optTodos 
               Height          =   285
               Left            =   150
               TabIndex        =   15
               Top             =   210
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   503
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
               Caption         =   "Todos"
               Value           =   -1
            End
            Begin Threed.SSOption optFarmacia 
               Height          =   285
               Left            =   1785
               TabIndex        =   16
               Top             =   210
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   503
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
               Caption         =   "Solo Farmacia"
            End
            Begin Threed.SSOption optServicios 
               Height          =   285
               Left            =   3615
               TabIndex        =   17
               Top             =   210
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   503
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
               Caption         =   "Sólo Servicios"
            End
         End
         Begin VB.Frame Frame 
            Height          =   2865
            Index           =   0
            Left            =   6330
            TabIndex        =   2
            Top             =   885
            Width           =   7665
            Begin VB.ComboBox cmbTrimestre 
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
               ItemData        =   "RpAtencionesTotales.frx":28B0
               Left            =   2250
               List            =   "RpAtencionesTotales.frx":28C0
               Style           =   2  'Dropdown List
               TabIndex        =   4
               Top             =   180
               Width           =   5355
            End
            Begin VB.ComboBox cmbTipoServicio 
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
               ItemData        =   "RpAtencionesTotales.frx":290D
               Left            =   2250
               List            =   "RpAtencionesTotales.frx":291D
               Style           =   2  'Dropdown List
               TabIndex        =   3
               Top             =   930
               Width           =   5355
            End
            Begin Threed.SSOption optEconXsexo 
               Height          =   285
               Left            =   75
               TabIndex        =   5
               Top             =   1455
               Width           =   2715
               _ExtentX        =   4789
               _ExtentY        =   503
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
               Caption         =   "Atenciones por Edad y Sexo"
               Value           =   -1
            End
            Begin Threed.SSOption optEconMorbilidad 
               Height          =   285
               Left            =   75
               TabIndex        =   6
               Top             =   1770
               Width           =   2325
               _ExtentX        =   4101
               _ExtentY        =   503
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
               Caption         =   "Morbilidad  (atenciones)"
            End
            Begin Threed.SSOption optEconXserviciosinterm 
               Height          =   285
               Left            =   75
               TabIndex        =   7
               Top             =   2100
               Width           =   3150
               _ExtentX        =   5556
               _ExtentY        =   503
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
               Caption         =   "Consumo en Servicios Intermedios"
            End
            Begin Threed.SSOption optEconRecetas 
               Height          =   285
               Left            =   75
               TabIndex        =   8
               Top             =   2415
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   503
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
               Caption         =   "Consumo de Recetas"
            End
            Begin MSDataListLib.DataCombo cmbFuenteFinanciamiento 
               Height          =   330
               Left            =   2250
               TabIndex        =   9
               Top             =   570
               Width           =   5355
               _ExtentX        =   9446
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
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "(usa Fecha Prescripción de RECETA)"
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
               Left            =   2295
               TabIndex        =   47
               Top             =   2460
               Width           =   2970
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "(usa Fecha: Cita/AltaMédicaHosp/AltaMédicaEmerg)"
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
               Left            =   3285
               TabIndex        =   46
               Top             =   2115
               Width           =   4350
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "(usa Fecha: Cita/AltaMédicaHosp/AltaMédicaEmerg)"
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
               Left            =   2415
               TabIndex        =   45
               Top             =   1815
               Width           =   4200
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Trimestre (alta médica)"
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
               TabIndex        =   13
               Top             =   270
               Width           =   1905
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Fuente Financiamiento"
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
               Top             =   630
               Width           =   1845
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Servicio"
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
               Top             =   1005
               Width           =   1035
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "(usa Fecha: Cita/AdmisiónHosp/AdmisiónEmerg)"
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
               Left            =   3030
               TabIndex        =   10
               Top             =   1485
               Width           =   3900
            End
         End
         Begin Threed.SSOption optAtenciones 
            Height          =   285
            Left            =   165
            TabIndex        =   18
            Top             =   690
            Width           =   5355
            _ExtentX        =   9446
            _ExtentY        =   503
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
            Caption         =   "Atenciones en CE, Hospitalización y Emergencia"
            Value           =   -1
         End
         Begin MSMask.MaskEdBox txtAnio 
            Height          =   315
            Left            =   630
            TabIndex        =   19
            Top             =   210
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####"
            PromptChar      =   "_"
         End
         Begin Threed.SSOption optPromDiasHosp 
            Height          =   285
            Left            =   165
            TabIndex        =   20
            Top             =   1050
            Width           =   5355
            _ExtentX        =   9446
            _ExtentY        =   503
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
            Caption         =   "Promedio de ESTANCIA de Hospitalización"
         End
         Begin Threed.SSOption optPromFactHosp 
            Height          =   285
            Left            =   165
            TabIndex        =   21
            Top             =   1380
            Width           =   5355
            _ExtentX        =   9446
            _ExtentY        =   503
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
            Caption         =   "Promedio de FACTURACION en Hospitalización"
         End
         Begin Threed.SSOption optNroVivosHosp 
            Height          =   285
            Left            =   165
            TabIndex        =   22
            Top             =   2250
            Width           =   5355
            _ExtentX        =   9446
            _ExtentY        =   503
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
            Caption         =   "Número de Egresos Vivos y Fallecidos en Hospitalización"
         End
         Begin Threed.SSOption optSISxDpto 
            Height          =   285
            Left            =   165
            TabIndex        =   23
            Top             =   2580
            Width           =   6150
            _ExtentX        =   10848
            _ExtentY        =   503
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
            Caption         =   "Sis,NoSis,SisSinSubsidioTotal/Referidos/EgresosYreingresos x Dptos"
         End
         Begin Threed.SSOption optPacientesYcpt 
            Height          =   285
            Left            =   165
            TabIndex        =   24
            Top             =   2910
            Width           =   6150
            _ExtentX        =   10848
            _ExtentY        =   503
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
            Caption         =   "Paciente y sus CPT atendidos"
         End
         Begin Threed.SSOption optReptrimestre 
            Height          =   285
            Left            =   6045
            TabIndex        =   25
            Top             =   645
            Width           =   6405
            _ExtentX        =   11298
            _ExtentY        =   503
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
            Caption         =   "Reportes Trimestrales"
         End
         Begin Threed.SSOption optEmegMenor24hr 
            Height          =   285
            Left            =   165
            TabIndex        =   64
            Top             =   3210
            Width           =   5925
            _ExtentX        =   10451
            _ExtentY        =   503
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
            Caption         =   "Pacientes en sala de Observación con estacia mayor o gual a 24 hr"
         End
         Begin Threed.SSOption emergReingresos24hr 
            Height          =   285
            Left            =   165
            TabIndex        =   65
            Top             =   3495
            Width           =   5925
            _ExtentX        =   10451
            _ExtentY        =   503
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
            Caption         =   "Reingresos a Emergencia <24 hr"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Año"
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
            Left            =   210
            TabIndex        =   26
            Top             =   240
            Width           =   330
         End
      End
      Begin Threed.SSOption optAtencionesAnuales 
         Height          =   285
         Left            =   120
         TabIndex        =   27
         Top             =   330
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   503
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
         Caption         =   "Atenciones ANUALES"
         Value           =   -1
      End
      Begin Threed.SSOption optMuertes 
         Height          =   285
         Left            =   120
         TabIndex        =   48
         Top             =   4575
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   503
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
         Caption         =   "Fallecidos"
      End
      Begin Threed.SSOption optXrangoFechas 
         Height          =   285
         Left            =   120
         TabIndex        =   60
         Top             =   6300
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   503
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
         Caption         =   "Por rango de Fechas"
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Huelga"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   32
      Top             =   285
      Visible         =   0   'False
      Width           =   6615
      Begin VB.TextBox txtExcel 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
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
         Height          =   555
         Left            =   1500
         MultiLine       =   -1  'True
         TabIndex        =   42
         Top             =   720
         Width           =   4635
      End
      Begin VB.ComboBox cmbAtencionTipo 
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
         ItemData        =   "RpAtencionesTotales.frx":2957
         Left            =   1500
         List            =   "RpAtencionesTotales.frx":2964
         TabIndex        =   41
         Top             =   390
         Width           =   4635
      End
      Begin VB.TextBox txtHospital 
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
         Left            =   1500
         TabIndex        =   40
         Top             =   1290
         Width           =   4635
      End
      Begin VB.Frame Frame4 
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
         Height          =   1695
         Left            =   150
         TabIndex        =   33
         Top             =   1680
         Width           =   6075
         Begin Threed.SSOption optAtencionesHistXanio 
            Height          =   285
            Left            =   120
            TabIndex        =   34
            Top             =   330
            Width           =   5355
            _ExtentX        =   9446
            _ExtentY        =   503
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
            Caption         =   "Usando datos (históricos) del AÑO ANTERIOR"
         End
         Begin Threed.SSOption AtencionesHistoricasSemanaAnt 
            Height          =   285
            Left            =   120
            TabIndex        =   35
            Top             =   690
            Width           =   5355
            _ExtentX        =   9446
            _ExtentY        =   503
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
            Caption         =   "Usando datos (históricos) de la SEMANA ANTERIOR"
            Value           =   -1
         End
         Begin MSMask.MaskEdBox txtFechaInicioH 
            Height          =   315
            Left            =   1425
            TabIndex        =   36
            Top             =   1170
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
         Begin MSMask.MaskEdBox txtFechaFinH 
            Height          =   315
            Left            =   3360
            TabIndex        =   37
            Top             =   1170
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
         Begin VB.Label Label4 
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
            Left            =   2865
            TabIndex        =   39
            Top             =   1215
            Width           =   435
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Días de Huelga"
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
            TabIndex        =   38
            Top             =   1215
            Width           =   1200
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
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
         TabIndex        =   44
         Top             =   390
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Hospital/Cs/Ps"
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
         TabIndex        =   43
         Top             =   1380
         Width           =   1140
      End
   End
   Begin Threed.SSOption SSOption2 
      Height          =   285
      Left            =   45
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   503
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
      Caption         =   "Atenciones Consulta Externa durante HUELGA"
   End
End
Attribute VB_Name = "RpAtencionesTotales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Atenciones por año
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim sMensaje As String
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim oRsFuenteFinaniamiento As New Recordset

Private Sub btnAceptar_Click()

If wxFranklin = "*" Then Exit Sub

        Me.MousePointer = 11
    
        Dim oRpt As New clAtencionesTotales
        Dim lcTitulo As String
        lcTitulo = ""
        If optXrangoFechas.Value = True Then
            If optDemoraAtencion.Value = True Then
               oRpt.TiempoEnAtenciones "0", 0, 0, optDemoraAtencion.Caption, CDate(Me.txtfCitaIni.Text), _
                                       CDate(Me.txtFcitaFin.Text)
            End If
        ElseIf optMuertes.Value = True Then
            If optMuertesHosp.Value = True Then
               lcTitulo = Trim(optMuertesHosp.Caption)
               oRpt.ProcesaMuerteHospitalizados lcTitulo, "(Fechas: " & txtFechaInicio.Text & " al " & txtFechaFin.Text & ")", _
                                                Me.txtFechaInicio.Text, Me.txtFechaFin.Text, 3, True, Me.hwnd
            ElseIf optMuerteEmergencia.Value = True Then
               lcTitulo = Trim(optMuerteEmergencia.Caption)
               oRpt.ProcesaMuerteHospitalizados lcTitulo, "(Fechas: " & txtFechaInicio.Text & " al " & txtFechaFin.Text & ")", _
                                                Me.txtFechaInicio.Text, Me.txtFechaFin.Text, 2, True, Me.hwnd
            ElseIf optMuerteHospME.Value = True Then
               lcTitulo = Trim(optMuerteHospME.Caption) & "  ( hasta " & txtEdad.Text & "  años)"
               oRpt.ProcesaMuerteHospitalizadosMenoresEdad lcTitulo, "(Fechas: " & txtFechaInicio.Text & " al " & txtFechaFin.Text & ")", _
                                                Me.txtFechaInicio.Text, Me.txtFechaFin.Text, 3, True, Val(txtEdad.Text), Me.hwnd
            End If
        ElseIf optAtencionesAnuales.Value = True Then
            If ValidaDatosObligatorios Then
            
                If optSISxDpto.Value = True Then
                    MsgBox "..este reporte se está mejorando..."
                    'oRpt.SISxDpto Me.hwnd, Me.txtAnio.Text, 1
                    Me.MousePointer = 1
                    Exit Sub
                ElseIf optEmegMenor24hr.Value = True Then
                    cmdObs
                    Exit Sub
                ElseIf emergReingresos24hr.Value = True Then
                    cmdReingresosEm
                    Exit Sub
                ElseIf optPacientesYcpt.Value = True Then
                    MsgBox "..este reporte se está mejorando..."
                    'oRpt.PacientesYcpt Me.hwnd, Me.txtAnio.Text
                    Me.MousePointer = 1
                    Exit Sub
                ElseIf Me.optPromFactHosp.Value = True Then
                   
                   lcTitulo = Trim(Me.optPromFactHosp.Caption) & " (" & IIf(Me.optTodos.Value = True, optTodos.Caption, _
                                                                    IIf(Me.optServicios.Value = True, optServicios.Caption, _
                                                                    Me.optFarmacia.Caption)) & ")"
                ElseIf optReptrimestre.Value = True And Me.optEconXsexo = True Then
                   lcTitulo = "Año: " & txtAnio.Text & " (" & cmbTrimestre.Text & ") (Fuente Financiamiento: " & _
                               cmbFuenteFinanciamiento.Text & ")" & _
                               IIf(cmbTipoServicio.ListIndex = 0, "", " (Tipo de Servicio: " & cmbTipoServicio.Text & ")")
                   oRpt.AtencionesXsexoEdad txtAnio.Text, cmbTrimestre.ListIndex, Val(cmbFuenteFinanciamiento.BoundText), _
                                            cmbTipoServicio.ListIndex, UCase(optEconXsexo.Caption), lcTitulo, True, Me.hwnd
                ElseIf optReptrimestre.Value = True And Me.optEconMorbilidad = True Then
                    lcTitulo = "Año: " & txtAnio.Text & " (" & cmbTrimestre.Text & ") (Fuente Financiamiento: " & _
                               cmbFuenteFinanciamiento.Text & ")" & _
                               IIf(cmbTipoServicio.ListIndex = 0, "", " (Tipo de Servicio: " & cmbTipoServicio.Text & ")")
                    oRpt.AtencionesMorbilidad txtAnio.Text, cmbTrimestre.ListIndex, Val(cmbFuenteFinanciamiento.BoundText), _
                                            cmbTipoServicio.ListIndex, UCase(optEconMorbilidad.Caption), lcTitulo, True, Me.hwnd
                ElseIf optReptrimestre.Value = True And Me.optEconXserviciosinterm.Value = True Then
                    lcTitulo = "Año: " & txtAnio.Text & " (" & cmbTrimestre.Text & ") (Fuente Financiamiento: " & _
                               cmbFuenteFinanciamiento.Text & ")" & _
                               IIf(cmbTipoServicio.ListIndex = 0, "", " (Tipo de Servicio: " & cmbTipoServicio.Text & ")")
                    oRpt.AtencionesServiciosIntermedios txtAnio.Text, cmbTrimestre.ListIndex, Val(cmbFuenteFinanciamiento.BoundText), _
                                            cmbTipoServicio.ListIndex, UCase(optEconXserviciosinterm.Caption), lcTitulo, True, Me.hwnd
                ElseIf Me.optReptrimestre.Value = True And optEconRecetas.Value = True Then
                    lcTitulo = "Año: " & txtAnio.Text & " (" & cmbTrimestre.Text & ") (Fuente Financiamiento: " & _
                               cmbFuenteFinanciamiento.Text & ")" & _
                               IIf(cmbTipoServicio.ListIndex = 0, "", " (Tipo de Servicio: " & cmbTipoServicio.Text & ")")
                    oRpt.AtencionesRecetas txtAnio.Text, cmbTrimestre.ListIndex, Val(cmbFuenteFinanciamiento.BoundText), _
                                            cmbTipoServicio.ListIndex, UCase(optEconRecetas.Caption), lcTitulo, True, Me.hwnd
                
                Else
                   oRpt.ProcesaReporte Me.txtAnio.Text, IIf(Me.optAtenciones.Value = True, 1, _
                                                    IIf(Me.optPromDiasHosp.Value = True, 2, _
                                                    IIf(Me.optPromFactHosp.Value = True, 3, 4))), _
                                                    IIf(Me.optTodos.Value = True, 1, IIf(Me.optServicios.Value = True, 2, 3)), _
                                                    IIf(Me.optAtenciones.Value = True, Me.optAtenciones.Caption, _
                                                    IIf(Me.optPromDiasHosp.Value = True, Me.optPromDiasHosp.Caption, _
                                                    IIf(Me.optPromFactHosp.Value = True, lcTitulo, _
                                                    Me.optNroVivosHosp.Caption))), Me.hwnd
                End If
            End If
        Else
            oRpt.ProcesaReporteParaHuelga cmbAtencionTipo.ListIndex, txtHospital.Text, _
                                           CDate(txtFechaInicioH.Text), CDate(txtFechaFinH.Text), _
                                           IIf(optAtencionesHistXanio.Value = True, 1, 2), Me.hwnd, _
                                           "Número de Atenciones en CE (" & Trim(Me.txtHospital.Text) & ")", _
                                           "(Tipo: " & Trim(cmbAtencionTipo.Text) & ") (Fechas: " & txtFechaInicioH.Text & _
                                           " " & txtFechaFinH.Text & ") (" & _
                                           IIf(optAtencionesHistXanio.Value = True, "Historicos AÑO ANTERIOR)", "Historicos SEMANA ANTERIOR)")

        End If
        Set oRpt = Nothing
        Me.MousePointer = 1
    
End Sub

Function ValidaDatosObligatorios() As Boolean
    sMensaje = ""
    If optAtencionesAnuales.Value = True Then
    Else
       If cmbAtencionTipo.Text = "" Then
          sMensaje = sMensaje & " Por favor elija el TIPO" & Chr(13)
       End If
       If txtHospital.Text = "" Then
          sMensaje = sMensaje & " Por favor ingrese el NOMBRE DEL HOSPITAL/CS/PS" & Chr(13)
       End If
        If CDate(Me.txtFechaInicioH.Text) > CDate(Me.txtFechaFinH.Text) Then
           MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, "Reporte"
           Exit Function
        End If
    End If

    If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       ValidaDatosObligatorios = False
    Else
       ValidaDatosObligatorios = True
    End If
End Function

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub

Private Sub cmbAtencionTipo_Change()
    cmbAtencionTipo_Click
End Sub

Private Sub cmbAtencionTipo_Click()
    Select Case cmbAtencionTipo.ListIndex
    Case 0
        txtExcel.Text = "Necesita archivos: excel2013atendidos.xls, excel2014atendidos.xls en c:\archivos de pr...\galenhos\archivos"
    Case 1
        txtExcel.Text = "Necesita archivo: MovimientoHC.xls en c:\archivos de pr...\galenhos\archivos"
    Case 2
        txtExcel.Text = "Necesita archivos: excel2014auditoria.xls en c:\archivos de pr...\galenhos\archivos"
    End Select

End Sub

Private Sub Form_Load()
    Me.txtAnio.Text = Year(Date)
    Me.txtFechaInicioH.Text = Format(Date, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
    Me.txtFechaFinH.Text = Format(Date, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
    '
    cmbTrimestre.ListIndex = 0
    Set oRsFuenteFinaniamiento = mo_ReglasFacturacion.FuentesFinanciamientoDevuelveTodosSegunFiltro("idFuenteFinanciamiento<>5")
    Set cmbFuenteFinanciamiento.RowSource = oRsFuenteFinaniamiento
    cmbFuenteFinanciamiento.ListField = "Descripcion"
    cmbFuenteFinanciamiento.BoundColumn = "idFuenteFinanciamiento"
    cmbFuenteFinanciamiento.BoundText = "1"
    cmbTipoServicio.ListIndex = 0
    '
    Me.txtFechaInicio.Text = SIGHEntidades.PrimerFechaDDMMYYDelMesActual()
    Me.txtFechaFin = SIGHEntidades.UltimaFechaDDMMYYDelMesActual()
    txtEdad.Text = "18"
    
    Me.txtfCitaIni.Text = "01/01/" & Me.txtAnio.Text
    Me.txtFcitaFin.Text = "31/12/" & Me.txtAnio.Text
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
    AdministrarKeyPreview KeyCode
End Sub


Private Sub cmdReingresosEm()
      Dim oRsTmp1 As New Recordset
      Dim oRsExcel As New Recordset
      Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
      Dim lcSql As String
      lcSql = "SELECT   dbo.Atenciones.* " & _
" FROM            dbo.Atenciones " & _
" WHERE        (dbo.Atenciones.IdTipoServicio = 2) AND (dbo.Atenciones.idEstadoAtencion <> 0) AND" & _
"                         year(dbo.Atenciones.FechaIngreso)=" & txtAnio.Text & _
" ORDER BY dbo.Atenciones.idPaciente,dbo.Atenciones.FechaIngreso,dbo.Atenciones.horaIngreso"
        oRsTmp1.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
        If oRsTmp1.RecordCount > 0 Then
            Dim lnMes As Integer, lnCantidad As Long, lnIdPaciente As Long
            With oRsExcel
                 .Fields.Append "mes", adInteger
                 .Fields.Append "mesNombre", adVarChar, 40
                 .Fields.Append "Cantidad", adInteger
                 .LockType = adLockOptimistic
                 .Open
                 For lnMes = 1 To 12
                      .AddNew
                      .Fields!Mes = lnMes
                      .Fields!mesNombre = SIGHEntidades.DevuelveNombreMes(lnMes)
                      .Fields!cantidad = 0
                      .Update
                 Next
            End With
            oRsTmp1.MoveFirst
            Dim lnNumero As Integer
            Dim ldFechaHoraInicial As String
            Do While Not oRsTmp1.EOF
                 lnMes = Month(oRsTmp1!fechaIngreso)
                 lnIdPaciente = oRsTmp1!idPaciente
                 If IsNull(oRsTmp1!fechaEgreso) Then
                   ldFechaHoraInicial = 0
                 Else
                    ldFechaHoraInicial = CDate(Format(oRsTmp1!fechaEgreso, "dd/mm/yyyy") & " " & oRsTmp1!horaEgreso)
                 End If
                 lnCantidad = 0
                 lnNumero = 1
                 Do While Not oRsTmp1.EOF And lnIdPaciente = oRsTmp1!idPaciente
                      If lnNumero > 1 Then
                          If IsDate(ldFechaHoraInicial) Then
                                If DateDiff("h", ldFechaHoraInicial, _
                                                 CDate(Format(oRsTmp1!fechaIngreso, "dd/mm/yyyy") & " " & oRsTmp1!horaIngreso)) < 24 Then
                                   lnCantidad = lnCantidad + 1
                                End If
                          End If
                          If IsNull(oRsTmp1!fechaEgreso) Then
                              ldFechaHoraInicial = 0
                          Else
                               ldFechaHoraInicial = CDate(Format(oRsTmp1!fechaEgreso, "dd/mm/yyyy") & " " & oRsTmp1!horaEgreso)
                          End If
                       End If
                       oRsTmp1.MoveNext
                       lnNumero = lnNumero + 1
                       If oRsTmp1.EOF Then
                         Exit Do
                       End If
                  Loop
                  If lnCantidad > 0 Then
                        oRsExcel.MoveFirst
                        oRsExcel.Find "mes=" & lnMes
                        oRsExcel!cantidad = oRsExcel!cantidad + lnCantidad
                        oRsExcel.Update
                  End If
            Loop
            If oRsExcel.RecordCount = 0 Then
               MsgBox "no hay datos"
            Else
               mo_AdminReportes.ExportarRecordSetAexcel oRsExcel, emergReingresos24hr.Caption, "Año: " & txtAnio.Text, "", Me.hwnd
            End If
       Else
            MsgBox "no hay datos"
       End If
      Set oRsTmp1 = Nothing
      Set oRsExcel = Nothing
      Set mo_AdminReportes = Nothing
      Me.MousePointer = 1
End Sub

Private Sub cmdObs()
      Dim oRsTmp1 As New Recordset
      Dim oRsExcel As New Recordset
      Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
      Dim lcSql As String
      lcSql = "SELECT   dbo.Atenciones.idAtencion,   dbo.Atenciones.FechaIngreso, dbo.AtencionesEstanciaHospitalaria.IdServicio, " & _
"                         dbo.AtencionesEstanciaHospitalaria.FechaOcupacion," & _
"                         dbo.AtencionesEstanciaHospitalaria.HoraOcupacion," & _
"                         dbo.AtencionesEstanciaHospitalaria.FechaDesocupacion," & _
"                         dbo.AtencionesEstanciaHospitalaria.HoraDesocupacion" & _
" FROM            dbo.Atenciones INNER JOIN" & _
"                         dbo.AtencionesEstanciaHospitalaria ON" & _
"                         dbo.Atenciones.IdAtencion = dbo.AtencionesEstanciaHospitalaria.IdAtencion" & _
" WHERE        (dbo.Atenciones.IdTipoServicio = 2) AND (dbo.Atenciones.idEstadoAtencion <> 0) AND" & _
"                         (dbo.AtencionesEstanciaHospitalaria.IdServicio IN (104, 118, 119)) AND" & _
"                         (NOT (dbo.AtencionesEstanciaHospitalaria.FechaDesocupacion IS NULL))" & _
"                         and year(dbo.Atenciones.FechaIngreso)=" & txtAnio.Text & _
" ORDER BY dbo.Atenciones.FechaIngreso"
        oRsTmp1.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
        If oRsTmp1.RecordCount > 0 Then
            Dim lnMes As Integer, lnCantidad As Long
            With oRsExcel
                 .Fields.Append "mes", adVarChar, 40
                 .Fields.Append "Cantidad", adInteger
                 .LockType = adLockOptimistic
                 .Open
            End With
            oRsTmp1.MoveFirst
            Do While Not oRsTmp1.EOF
                 lnMes = Month(oRsTmp1!fechaIngreso)
                 lnCantidad = 0
                 Do While Not oRsTmp1.EOF And lnMes = Month(oRsTmp1!fechaIngreso)
                      If DateDiff("h", Format(oRsTmp1!fechaOcupacion, "dd/mm/yyyy") & " " & oRsTmp1!horaOcupacion, _
                                    Format(oRsTmp1!fechaDesocupacion, "dd/mm/yyyy") & " " & oRsTmp1!horaDesocupacion) >= 24 Then
                                    lnCantidad = lnCantidad + 1
                       End If
                       oRsTmp1.MoveNext
                       If oRsTmp1.EOF Then
                         Exit Do
                       End If
                  Loop
                  If lnCantidad > 0 Then
                        oRsExcel.AddNew
                        oRsExcel!Mes = SIGHEntidades.DevuelveNombreMes(lnMes)
                        oRsExcel!cantidad = lnCantidad
                        oRsExcel.Update
                  End If
            Loop
            If oRsExcel.RecordCount = 0 Then
               MsgBox "no hay datos"
            Else
               mo_AdminReportes.ExportarRecordSetAexcel oRsExcel, optEmegMenor24hr.Caption, "Año: " & txtAnio.Text, "", Me.hwnd
            End If
       Else
            MsgBox "no hay datos"
       End If
      Set oRsTmp1 = Nothing
      Set oRsExcel = Nothing
      Set mo_AdminReportes = Nothing
      Me.MousePointer = 1

End Sub

