VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmMantenimientoHIS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar hoja HIS"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12915
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMantenimientoHIS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   12915
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDetalleAtencion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   0
      TabIndex        =   21
      Top             =   4080
      Width           =   12855
      Begin VB.TextBox txtNroRegistro 
         Alignment       =   2  'Center
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
         MaxLength       =   2
         TabIndex        =   66
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
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
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   28
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtPeso 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00;(0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
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
         Left            =   6480
         MaxLength       =   6
         TabIndex        =   40
         Top             =   1080
         Width           =   855
      End
      Begin VB.ComboBox cmbTipoEdad 
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
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmdIngresarDiagnosticos 
         Caption         =   "+Dx"
         Height          =   915
         Left            =   12120
         TabIndex        =   44
         Top             =   480
         Width           =   615
      End
      Begin VB.ComboBox cmbFinanciador 
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
         Left            =   8520
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox cmbTipoDocumento 
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
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   480
         Width           =   2295
      End
      Begin VB.ComboBox cmbEstadoFrenteServicio 
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
         Left            =   10200
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   1080
         Width           =   1935
      End
      Begin VB.ComboBox cmbEstadoFrenteEstablecimiento 
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
         Left            =   8280
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   1080
         Width           =   1935
      End
      Begin VB.ComboBox cmbSexo 
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
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtOrdenFamiliar 
         Alignment       =   2  'Center
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
         Left            =   7800
         MaxLength       =   2
         TabIndex        =   33
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtDistritoProcedencia 
         Alignment       =   2  'Center
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
         TabIndex        =   36
         Top             =   1080
         Width           =   3135
      End
      Begin VB.ComboBox cmbEtnia 
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
         Left            =   10320
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   480
         Width           =   1755
      End
      Begin VB.TextBox txtNroDocumento 
         Alignment       =   2  'Center
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
         Left            =   6360
         MaxLength       =   12
         TabIndex        =   32
         ToolTipText     =   "Presione ENTER para Buscar en la Base de Datos Local"
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtNacionalidad 
         Alignment       =   2  'Center
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
         Left            =   3360
         MaxLength       =   3
         TabIndex        =   30
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtNroHC_FF_COD 
         Alignment       =   2  'Center
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
         Left            =   1920
         MaxLength       =   11
         TabIndex        =   29
         Top             =   480
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtTalla 
         Height          =   315
         Left            =   7320
         TabIndex        =   41
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         AutoTab         =   -1  'True
         HideSelection   =   0   'False
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin UltraGrid.SSUltraGrid ugvDetalleDiagnosticos 
         Height          =   2655
         Left            =   120
         TabIndex        =   45
         Top             =   1440
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   4683
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108868
         MaxColScrollRegions=   50
         MaxRowScrollRegions=   50
         Caption         =   "Detalle de Diagnósticos"
      End
      Begin MSMask.MaskEdBox txtEdad 
         Height          =   315
         Left            =   3240
         TabIndex        =   37
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         _Version        =   393216
         AutoTab         =   -1  'True
         HideSelection   =   0   'False
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Caption         =   "Nro Registro"
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
         TabIndex        =   67
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Talla(Cm)"
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
         Index           =   13
         Left            =   7320
         TabIndex        =   59
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "Peso (Kg)"
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
         Index           =   12
         Left            =   6480
         TabIndex        =   58
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "Tipo Documento"
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
         Index           =   10
         Left            =   4080
         TabIndex        =   54
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label13 
         Caption         =   "Edición de Atención - Diagnósticos"
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
         TabIndex        =   22
         Top             =   0
         Width           =   2895
      End
      Begin VB.Label Label15 
         Caption         =   "Servicio"
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
         Index           =   9
         Left            =   10200
         TabIndex        =   52
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Establecimiento"
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
         Index           =   8
         Left            =   8280
         TabIndex        =   51
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label15 
         Caption         =   "Sexo"
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
         Index           =   7
         Left            =   4680
         TabIndex        =   50
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Edad"
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
         Index           =   6
         Left            =   3240
         TabIndex        =   49
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "T. Edad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   3720
         TabIndex        =   48
         Top             =   840
         Width           =   840
      End
      Begin VB.Label Label15 
         Caption         =   "Procedencia"
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
         Index           =   4
         Left            =   120
         TabIndex        =   47
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Etnia"
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
         Index           =   3
         Left            =   10320
         TabIndex        =   46
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Financiador"
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
         Index           =   2
         Left            =   8520
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Nº Hjo"
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
         Index           =   1
         Left            =   7800
         TabIndex        =   26
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Nro Documento"
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
         Index           =   0
         Left            =   6360
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Nac."
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
         Left            =   3480
         TabIndex        =   55
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "HC-FF-COD ACT."
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
         Left            =   1920
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Día"
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
         Left            =   1320
         TabIndex        =   23
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   0
      TabIndex        =   57
      Top             =   8760
      Width           =   12855
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar"
         DisabledPicture =   "frmMantenimientoHIS.frx":000C
         DownPicture     =   "frmMantenimientoHIS.frx":046C
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
         Left            =   4680
         Picture         =   "frmMantenimientoHIS.frx":08E1
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   120
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Salir (ESC)"
         DisabledPicture =   "frmMantenimientoHIS.frx":0D56
         DownPicture     =   "frmMantenimientoHIS.frx":121A
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
         Left            =   6240
         Picture         =   "frmMantenimientoHIS.frx":1706
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   120
         Width           =   1365
      End
      Begin Threed.SSCommand btnAgregarHoja 
         Height          =   705
         Left            =   10800
         TabIndex        =   71
         ToolTipText     =   "Agregar visita"
         Top             =   120
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   1244
         _Version        =   262144
         CaptionStyle    =   1
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmMantenimientoHIS.frx":1BF2
         Caption         =   "Agregar Nueva Hoja al Lote HIS"
         PictureAlignment=   9
      End
   End
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
      Height          =   1575
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   12855
      Begin VB.ComboBox cmbMes 
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
         Left            =   10440
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtPagRestante 
         Alignment       =   2  'Center
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
         Left            =   9000
         MaxLength       =   3
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtUltimaPaginaLoteActiva 
         Alignment       =   2  'Center
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
         Left            =   7560
         MaxLength       =   2
         TabIndex        =   3
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtLote 
         Alignment       =   2  'Center
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
         MaxLength       =   3
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtNroPaginas 
         Alignment       =   2  'Center
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
         Left            =   6120
         MaxLength       =   3
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtUbigeoDist 
         Alignment       =   2  'Center
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
         TabIndex        =   11
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtUbigeoEstablecimiento 
         Alignment       =   2  'Center
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
         Left            =   1320
         TabIndex        =   12
         Top             =   480
         Width           =   3345
      End
      Begin VB.TextBox txtCodigoEstadistico 
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
         Left            =   8760
         TabIndex        =   20
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox txtResponsable 
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
         Left            =   4680
         TabIndex        =   19
         Top             =   1080
         Width           =   3975
      End
      Begin VB.ComboBox cmbServicioCodigo 
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1080
         Width           =   2775
      End
      Begin VB.ComboBox cmbTurno 
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
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1080
         Width           =   1815
      End
      Begin MSMask.MaskEdBox txtfechaAnio 
         Height          =   330
         Left            =   12000
         TabIndex        =   68
         Top             =   480
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   582
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
      Begin VB.Label Label17 
         Caption         =   "Mes"
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
         Left            =   10440
         TabIndex        =   65
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "Hojas Libres"
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
         Left            =   9000
         TabIndex        =   62
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label5 
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
         Height          =   255
         Left            =   12000
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Lote"
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
         Index           =   0
         Left            =   4920
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Total Hojas"
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
         Index           =   1
         Left            =   6120
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Hoja Actual"
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
         Left            =   7560
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label9 
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
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Establecimiento"
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
         Left            =   1320
         TabIndex        =   6
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "Responsable Digitación"
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
         Index           =   0
         Left            =   8880
         TabIndex        =   16
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label Label11 
         Caption         =   "Responsable Atención"
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
         Left            =   4695
         TabIndex        =   15
         Top             =   840
         Width           =   3960
      End
      Begin VB.Label Label10 
         Caption         =   "Servicio"
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
         Left            =   1930
         TabIndex        =   14
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Turno"
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
         TabIndex        =   13
         Top             =   840
         Width           =   615
      End
   End
   Begin UltraGrid.SSUltraGrid ugvResumenHIS 
      Height          =   2295
      Left            =   30
      TabIndex        =   69
      Top             =   1680
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   4048
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108868
      MaxColScrollRegions=   50
      MaxRowScrollRegions=   50
      Caption         =   "Detalle de Atenciones"
   End
   Begin UltraGrid.SSUltraGrid ugvResumenDiagnosticos 
      Height          =   2295
      Left            =   8760
      TabIndex        =   70
      Top             =   1680
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4048
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108868
      MaxColScrollRegions=   50
      MaxRowScrollRegions=   50
      Caption         =   "Diagnósticos"
   End
   Begin VB.Label lblMensajeDNIBusqueda 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4080
      TabIndex        =   63
      Top             =   8520
      Width           =   8775
   End
   Begin VB.Label Label15 
      Caption         =   $"frmMantenimientoHIS.frx":4B7E
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   0
      TabIndex        =   61
      Top             =   8280
      Width           =   11895
   End
   Begin VB.Label Label15 
      Caption         =   "(F11) Consultar Listas - (F12) Guardar Atención"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   0
      TabIndex        =   56
      Top             =   8520
      Width           =   11895
   End
End
Attribute VB_Name = "frmMantenimientoHIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Interfaz grafica en donde se ingresaran las atenciones del Pdron Nominal.
'        Programado por: Cachay F
'        Fecha: Febrero 2014
'
'------------------------------------------------------------------------------------

Option Explicit
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim mo_cmbTurno As New SIGHEntidades.ListaDespleglable
Dim mo_cmbServicioCodigo As New SIGHEntidades.ListaDespleglable
Dim mo_cmbTipoDocumento As New SIGHEntidades.ListaDespleglable
Dim mo_cmbFinanciador As New SIGHEntidades.ListaDespleglable
Dim mo_cmbEtnia As New SIGHEntidades.ListaDespleglable
Dim mo_cmbTipoEdad As New SIGHEntidades.ListaDespleglable
Dim mo_cmbSexo As New SIGHEntidades.ListaDespleglable
Dim mo_cmbEstadoFrenteEstablecimiento As New SIGHEntidades.ListaDespleglable
Dim mo_cmbEstadoFrenteServicio As New SIGHEntidades.ListaDespleglable
Dim mo_cmbMes As New SIGHEntidades.ListaDespleglable
Dim mb_PrimerIngresoCabeceraAtencion As Boolean
Public Event SePresionoTeclaEspecial(KeyCode As Integer)
'--------------------------- Variables de manejo de negocio -------------------------------
Dim oCabeceraAtencion As New DOHIS_Cabecera             'Contiene los datos de la cabecera de atencion
Dim mo_ReglasHIS As New SIGHNegocios.ReglasHISGalenos   'Representa la Capa de Negocios del Modulo HIS GalenHos
'Dim mo_DatosParametro As New SIGHDatos.Parametros       'Representa la fecha y hora del servidor
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim oRcs_DetalleAtencion As New Recordset               'Representa el detalle de las Atencion
Dim oRcs_DetalleAtencionTemp As New Recordset
Dim oRcs_Diagnosticos As New Recordset                  'Representa el detalle de Diagnosticos de la Atencion
Dim oRcs_DiagnosticosTemp As New Recordset              'Representa el detalle de Diagnosticos para una Atencion, solo existe por Atencion.
Dim mr_ReglasHIS As New SIGHNegocios.ReglasHISGalenos
Dim ml_IdCabeceraHIS As Long
Dim ml_IdLote As Long
Dim ml_IdUsuario As Long
Dim mo_lcNombrePc As String
Dim mo_lnIdTablaLISTBARITEMS  As Long
Dim mi_Opcion As sghOpciones
Dim mi_BotonPresionado As sghBotonDetallePresionado

'Datos de Inicio de Formulario
Dim ml_IdDepartamentoActual As Long: Dim ms_NombreDepActual As String
Dim ml_IdProvinciaActual As Long: Dim ms_NombreProvActual As String
Dim ml_IdDistritoActual As Long: Dim ms_NombreDistrActual As String
Dim ml_IdEstablecimientoActual As Long: Dim ms_CodigoEstablecimiento As String: Dim ms_NombreEstablecimientoActual As String
Dim ml_IdEstablecimiento As Long
Dim mb_SeleccionoLote As Boolean: Dim mb_SeleccionoHoja As Boolean: Dim mb_SeleccionoMedico As Boolean
Dim mb_PesoTallaHabilitados As Boolean

Dim mo_LoteActual As New DOHIS_Lotes
Dim ml_CodigoResponsableDigitacion As Long: Dim ms_NombreRespDigitacion As String

'Datos de Proceso
Dim IdAtencion As Long
Dim IdDiagnostico As Long
Dim IdAtencionMax As Long
Dim IdDiagnosticoMax As Long
Dim mb_FaltaGrabarAtencion As Boolean
Dim ml_IdPacienteGalenHos As Long

Dim IdTipoActividad As Integer
Dim IdDetalleDiagnostico As Integer
Dim ml_IdMedicoResponsable As Long
Dim ml_IdDistritoAtencion As Long
Dim ml_IdNacionalidadAtencion As Long
Const mi_CantidadMaxDiagnosticos As Integer = 6

'Datos de Cuadros de Dialogo
Dim IdCodigoActividad As Long
Dim IdCodigoNacionalidad As Long

'Colores de Indicacion
Const MI_IDNACIONALIDAD As Integer = 166
Const MS_NOMBRE_NAC As String = "PER"
Const ml_ColorCorrecto As Long = &HFFFFFF
Const ml_ColorError As Long = &HFF6347
Const ml_ColorMensaje As Long = &HFCD33E

'========================================== PROPIEDADES ====================================
'Propiedades de Modulo
Property Let CabeceraAtencion(oValue As DOHIS_Cabecera)
   Set oCabeceraAtencion = oValue
End Property

Property Get CabeceraAtencion() As DOHIS_Cabecera
   CabeceraAtencion = oCabeceraAtencion
End Property

Property Let IdHisCabecera(lValue As Long)
   ml_IdCabeceraHIS = lValue
End Property

Property Get IdHisCabecera() As Long
   IdHisCabecera = ml_IdCabeceraHIS
End Property

'Propiedades del Sistema
Property Let IdLoteHIS(lValue As Long)
   ml_IdLote = lValue
End Property

Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
End Property

Property Let IdEstablecimiento(lValue As Long)
    ml_IdEstablecimiento = lValue
End Property
Property Get IdEstablecimiento() As Long
    IdEstablecimiento = ml_IdEstablecimiento
End Property

Property Let Opcion(lValue As sghOpciones)
   mi_Opcion = lValue
End Property

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property

Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Property Let BotonPresionado(oValue As sghBotonDetallePresionado)
   mi_BotonPresionado = oValue
End Property

Property Get BotonPresionado() As sghBotonDetallePresionado
   BotonPresionado = mi_BotonPresionado
End Property

Private Sub btnDatosExtras_Click()
    Dim mo_DetallePadronInicial As New frmPadronNominal
    mo_DetallePadronInicial.IdUsuario = ml_IdUsuario
    'mo_DetalleDia.IdMes = CLng(Me.txtMes.Text)
    'mo_DetalleDia.IdAnio = CLng(txtFechaAnio.Text)
    mo_DetallePadronInicial.MostrarFormulario
End Sub

Private Sub btnAgregarHoja_Click()
    If oRcs_DetalleAtencion.RecordCount = 0 Then
        MsgBox "Para poder registrar otra hoja, esta hoja debe tener por lo menos un registro de atención", vbInformation, Me.Caption
        Exit Sub
    End If
    
    Dim orsTemp As Recordset
    Set orsTemp = ListadoHojasLibre
    
    If orsTemp.RecordCount = 0 Then
        MsgBox "Todas las hojas del lote ya fueron creadas", vbInformation, Me.Caption
        Exit Sub
    End If
    
    If orsTemp.RecordCount > 0 Then
        ActualizacionCerrarFormulario 'Guarda la hoja para activar la siguiente
        mi_Opcion = sghAgregar
'        With oCabeceraAtencion
'            .IdHisCabecera = 0
'            .IdHisLote = ml_IdLote
'            .IdEstablecimiento = ml_IdEstablecimientoActual
'            .IdServicio = mo_cmbServicioCodigo.BoundText
'            .IdMedico = ml_IdMedicoResponsable
'            .FechaCreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL
'            .IdUsuario = ml_IdUsuario
'            .NroFormato = mo_ReglasHIS.ObtenerDatosNroFormatoLibre(CInt(Me.txtFechaAnio.Text), ml_IdEstablecimientoActual)
'            .NroHojaHis = Val(Me.txtUltimaPaginaLoteActiva.Text)
'            .IdEstadoHis = 1
'            .IdTurno = Val(Me.cmbTurno.ItemData(Me.cmbTurno.ListIndex))
'        End With
'        ml_IdCabeceraHIS = mo_ReglasHIS.IngresarHojaHIS(oCabeceraAtencion)
        
        orsTemp.MoveFirst
        txtUltimaPaginaLoteActiva.Text = orsTemp.Fields!IdHoja
        txtPagRestante.Text = Val(txtPagRestante.Text) - 1
        
        If oRcs_DetalleAtencion.RecordCount <> 0 Then
            oRcs_DetalleAtencion.MoveFirst
            Do While Not oRcs_DetalleAtencion.EOF
                oRcs_DetalleAtencion.Delete
                oRcs_DetalleAtencion.MoveNext
            Loop
        End If
        
        Dim oRcs_DiagnosticosTemp1 As Recordset
        Dim oRcs_DiagnosticosTemp2 As New Recordset
        Set Me.ugvResumenDiagnosticos.DataSource = oRcs_DiagnosticosTemp2
        
        mb_SeleccionoLote = True
        mb_SeleccionoHoja = True
        mb_SeleccionoMedico = False
        
        HabilitaDeshabilitarPorOpcion
        
        'ASIGNAR LOS VALORES POR DEFECTO DEL REGISTRO DE ATENCION
        If mi_Opcion = sghAgregar Then
            ControlesAtencionPorDefecto
            SeleccionaNroRegistroLibre
        End If
        ml_IdCabeceraHIS = 0
        mo_Formulario.HabilitarDeshabilitar Me.txtResponsable, True
        Me.txtResponsable.Text = ""
        Me.txtResponsable.SetFocus
        
        Me.btnAgregarHoja.Enabled = False
    End If
End Sub

Public Function ListadoHojasLibre() As Recordset
    Dim lnIndice As Integer
    Dim HojaUsada As Boolean
    Dim oRcs_HojasLibres As New Recordset
    Dim oRcsTemp As Recordset
    'Para cargar los datos de una consulta
    Set ListadoHojasLibre = Nothing
    With oRcs_HojasLibres
        .Fields.Append "IdHoja", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "Hoja", adVarChar, 30, adFldIsNullable + adFldUpdatable
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
    Set oRcsTemp = mr_ReglasHIS.His_ConsultarHojasRegistradas(ml_IdEstablecimiento, ml_IdLote)
    For lnIndice = 1 To Val(txtNroPaginas.Text)
        HojaUsada = False
        If oRcsTemp.RecordCount <> 0 Then
            oRcsTemp.MoveFirst
            Do While Not oRcsTemp.EOF
                If lnIndice = Int(oRcsTemp!NroHojaHis) Then
                    HojaUsada = True
                End If
                oRcsTemp.MoveNext
            Loop
        End If
        If HojaUsada = False Then
            With oRcs_HojasLibres
                .AddNew
                .Fields!IdHoja = lnIndice
                .Fields!Hoja = "Hoja Nº " & lnIndice
                .Update
            End With
        End If
    Next
    If oRcs_HojasLibres.RecordCount > 0 Then oRcs_HojasLibres.MoveFirst
    Set ListadoHojasLibre = oRcs_HojasLibres
End Function

Private Sub cmbServicioCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtResponsable
End Sub

Private Sub cmbTurno_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbServicioCodigo
End Sub

Private Sub cmdIngresarDiagnosticos_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            mo_Teclado.RealizarNavegacion KeyCode, Me.ugvDetalleDiagnosticos
        Case vbKeyF12   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
            GrabaAtencionDiagnosticos
        Case vbKeyF10
            Me.ugvDetalleDiagnosticos.Update
            AdicionDiagnostico
        Case vbKeyF7    'CANCELA EDICION DE ATENCION
            CancelaEdicionAtencion
        Case vbKeyEscape
            If MsgBox("Desea salir del registro de la Hoja HIS?", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
                btnCancelar_Click
            End If
    End Select
End Sub

'========================================== EVENTOS ========================================
Private Sub Form_Initialize()
    Set mo_cmbTurno.MiComboBox = Me.cmbTurno
    Set mo_cmbServicioCodigo.MiComboBox = Me.cmbServicioCodigo
    Set mo_cmbMes.MiComboBox = Me.cmbMes
    Set mo_cmbTipoDocumento.MiComboBox = Me.cmbTipoDocumento
    Set mo_cmbFinanciador.MiComboBox = Me.cmbFinanciador
    Set mo_cmbEtnia.MiComboBox = Me.cmbEtnia
    Set mo_cmbTipoEdad.MiComboBox = Me.cmbTipoEdad
    Set mo_cmbSexo.MiComboBox = Me.cmbSexo
    Set mo_cmbEstadoFrenteEstablecimiento.MiComboBox = Me.cmbEstadoFrenteEstablecimiento
    Set mo_cmbEstadoFrenteServicio.MiComboBox = Me.cmbEstadoFrenteServicio
End Sub

Private Sub Form_Load()
    CrearTablasTemp
    CargarCombosCabecera
    CargarDatosAlFormulario
    CargarCombosDetalle
    
    mo_Formulario.HabilitarDeshabilitar Me.txtUbigeoDist, False
    mo_Formulario.HabilitarDeshabilitar Me.txtUbigeoEstablecimiento, False
    mo_Formulario.HabilitarDeshabilitar Me.cmbMes, False
    mo_Formulario.HabilitarDeshabilitar Me.txtFechaAnio, False
    mo_Formulario.HabilitarDeshabilitar Me.txtNroPaginas, False
    mo_Formulario.HabilitarDeshabilitar Me.txtPagRestante, False
    mo_Formulario.HabilitarDeshabilitar Me.txtCodigoEstadistico, False
    mo_Formulario.HabilitarDeshabilitar cmbTurno, False
    mo_Formulario.HabilitarDeshabilitar cmbServicioCodigo, False
    txtNacionalidad.Locked = True
    
    HabilitaDeshabilitarPorOpcion
    
    mo_Apariencia.ConfigurarFilasBiColores Me.ugvResumenHIS, SIGHEntidades.GrillaConFilasBicolor
    mo_Apariencia.ConfigurarFilasBiColores Me.ugvResumenDiagnosticos, SIGHEntidades.GrillaConFilasBicolor
    mo_Apariencia.ConfigurarFilasBiColores Me.ugvDetalleDiagnosticos, SIGHEntidades.GrillaConFilasBicolor
    
    'ASIGNAR LOS VALORES POR DEFECTO DEL REGISTRO DE ATENCION
    If mi_Opcion = sghAgregar Or mi_Opcion = sghModificar Then
        ControlesAtencionPorDefecto
        SeleccionaNroRegistroLibre
    End If
    
    btnAgregarHoja.Enabled = False
    If mi_Opcion = sghModificar Then Me.btnAgregarHoja.Enabled = True
'    If mi_Opcion = sghModificar Then mo_Teclado.RealizarNavegacion KeyCode, txtDia
End Sub

Public Sub HabilitaDeshabilitarPorOpcion()
    Select Case mi_Opcion
    Case sghAgregar
        Me.Caption = "Agregar Hoja HIS"
        mo_Formulario.HabilitarDeshabilitar Me.cmdIngresarDiagnosticos, False
        mb_PesoTallaHabilitados = False
        mb_PrimerIngresoCabeceraAtencion = True
        mo_Formulario.HabilitarDeshabilitar Me.txtNroRegistro, False
        mo_Formulario.HabilitarDeshabilitar Me.txtDia, False
        mo_Formulario.HabilitarDeshabilitar Me.txtNroHC_FF_COD, False
        mo_Formulario.HabilitarDeshabilitar Me.txtNacionalidad, False
        mo_Formulario.HabilitarDeshabilitar Me.cmbTipoDocumento, False
        mo_Formulario.HabilitarDeshabilitar Me.txtNroPaginas, False
        mo_Formulario.HabilitarDeshabilitar Me.txtNroDocumento, False
        mo_Formulario.HabilitarDeshabilitar Me.txtOrdenFamiliar, False
        mo_Formulario.HabilitarDeshabilitar Me.cmbFinanciador, False
        mo_Formulario.HabilitarDeshabilitar Me.cmbEtnia, False
        mo_Formulario.HabilitarDeshabilitar Me.txtDistritoProcedencia, False
        mo_Formulario.HabilitarDeshabilitar Me.txtEdad, False
        mo_Formulario.HabilitarDeshabilitar Me.cmbTipoEdad, False
        mo_Formulario.HabilitarDeshabilitar Me.cmbSexo, False
        mo_Formulario.HabilitarDeshabilitar Me.txtPeso, False
        mo_Formulario.HabilitarDeshabilitar Me.txtTalla, False
        mo_Formulario.HabilitarDeshabilitar Me.cmbEstadoFrenteEstablecimiento, False
        mo_Formulario.HabilitarDeshabilitar Me.cmbEstadoFrenteServicio, False
            
    Case sghModificar, sghConsultar, sghEliminar
        mo_Formulario.HabilitarDeshabilitar Me.txtLote, False
        mo_Formulario.HabilitarDeshabilitar Me.txtUltimaPaginaLoteActiva, False
        mo_Formulario.HabilitarDeshabilitar Me.cmbTurno, False
        mo_Formulario.HabilitarDeshabilitar Me.cmbServicioCodigo, False
        mo_Formulario.HabilitarDeshabilitar Me.txtResponsable, False
        mo_Formulario.HabilitarDeshabilitar Me.cmdIngresarDiagnosticos, True
           
        If mi_Opcion = sghModificar Then
            Me.Caption = "Modificar Hoja HIS"
            If lcBuscaParametro.SeleccionaFilaParametro(329) = "S" Then
                 mo_Formulario.HabilitarDeshabilitar Me.txtPeso, True
                 mo_Formulario.HabilitarDeshabilitar Me.txtTalla, True
            Else
                 mo_Formulario.HabilitarDeshabilitar Me.txtPeso, False
                 mo_Formulario.HabilitarDeshabilitar Me.txtTalla, False
            End If
        ElseIf mi_Opcion = sghConsultar Then
            Me.Caption = "Consultar Hoja HIS"
            Me.cmdIngresarDiagnosticos.Enabled = False
            BloquearControlesAtencion
        ElseIf mi_Opcion = sghEliminar Then
            Me.Caption = "Eliminar Hoja HIS"
            Me.btnAceptar.Visible = True
            Me.cmdIngresarDiagnosticos.Enabled = False
            BloquearControlesAtencion
        End If
        Me.ugvResumenHIS.PerformAction ssKeyActionFirstCellInRow
        Me.ugvResumenHIS.PerformAction ssKeyActionUndoCell
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ActualizacionCerrarFormulario
End Sub

Public Sub ActualizacionCerrarFormulario()
    Dim lnTotalReg As Long
    If mi_Opcion = sghAgregar Then
        If oRcs_DetalleAtencion.RecordCount = 0 Then
            If ml_IdCabeceraHIS <> 0 Then
                With oCabeceraAtencion
                    .IdHisCabecera = ml_IdCabeceraHIS
                End With
                ml_IdCabeceraHIS = mo_ReglasHIS.EliminarHojaHIS(oCabeceraAtencion)
            End If
        Else
            lnTotalReg = oRcs_DetalleAtencion.RecordCount
            If ml_IdCabeceraHIS <> 0 Then
                With oCabeceraAtencion
                    .IdHisCabecera = ml_IdCabeceraHIS
                    .NroFormato = lnTotalReg
                End With
                ml_IdCabeceraHIS = mo_ReglasHIS.ModificarHojaHIS(oCabeceraAtencion)
            End If
        End If
    Else
        lnTotalReg = oRcs_DetalleAtencion.RecordCount
        If ml_IdCabeceraHIS <> 0 Then
            With oCabeceraAtencion
                .IdHisCabecera = ml_IdCabeceraHIS
                .NroFormato = lnTotalReg
            End With
        End If
        ml_IdCabeceraHIS = mo_ReglasHIS.ModificarHojaHIS(oCabeceraAtencion)
    End If
End Sub

Private Sub txtDia_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF11
        If Trim(Me.txtLote.Text) = "" Then
            MsgBox "No se digitó el lote", vbInformation, Me.Caption
            Exit Sub
        End If
        If Trim(Me.txtResponsable.Text) = "" Then
            MsgBox "No se digitó el responsable", vbInformation, Me.Caption
            Exit Sub
        End If
        Dim mo_DetalleDia As New frmDetalleDia
        mo_DetalleDia.IdEstablecimiento = ml_IdEstablecimientoActual
        mo_DetalleDia.IdServicio = mo_cmbServicioCodigo.BoundText
        mo_DetalleDia.IdMedicoResponsable = ml_IdMedicoResponsable
        mo_DetalleDia.IdMes = mo_cmbMes.BoundText
        mo_DetalleDia.IdAnio = CLng(txtFechaAnio.Text)
        mo_DetalleDia.IdTurno = mo_cmbTurno.BoundText
        mo_DetalleDia.MostrarFormulario
        If mo_DetalleDia.BotonPresionado = sghAceptar Then
            Me.txtDia.Text = mo_DetalleDia.IdDia
            mo_Teclado.RealizarNavegacion KeyCode, txtNroHC_FF_COD
        Else
             Me.txtDia.SetFocus
        End If
        Set mo_DetalleDia = Nothing
    Case vbKeyReturn
        mo_Teclado.RealizarNavegacion KeyCode, Me.txtNroHC_FF_COD
    Case vbKeyF12   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
        GrabaAtencionDiagnosticos
    Case vbKeyF7    'CANCELA EDICION DE ATENCION
        CancelaEdicionAtencion
    Case vbKeyEscape
        If MsgBox("Desea salir del registro de la Hoja HIS?", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
            btnCancelar_Click
        End If
    End Select
End Sub

Private Sub txtDia_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48) Or KeyAscii > 57 Then
        If KeyAscii = 8 Then
            KeyAscii = 8
        Else
            KeyAscii = 1
        End If
    End If
End Sub

Private Sub txtEdad_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48) Or KeyAscii > 57 Then
        If KeyAscii = 8 Then
            KeyAscii = 8
        Else
            KeyAscii = 1
        End If
    End If
End Sub



'Activar Busqueda de Lotes para el periodo actual
Private Sub txtLote_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF11
            Dim mo_DetalleLotes As New frmDetalleLotes
            mo_DetalleLotes.IdEstablecimiento = ml_IdEstablecimientoActual
            mo_DetalleLotes.MostrarFormulario
            If mo_DetalleLotes.BotonPresionado = sghAceptar Then
                ml_IdLote = mo_DetalleLotes.IdLote
                txtFechaAnio.Text = mo_DetalleLotes.Anio
                mo_cmbMes.BoundText = mo_DetalleLotes.IdMes
                txtLote.Text = mo_DetalleLotes.Lote
                txtNroPaginas.Text = mo_DetalleLotes.NumeroPaginas
                Me.txtPagRestante.Text = mo_DetalleLotes.NumeroPaginas - mr_ReglasHIS.His_ConsultarHojasRegistradas(ml_IdEstablecimientoActual, ml_IdLote).RecordCount - 1

                If Val(mo_DetalleLotes.NumeroPaginas) - Val(mr_ReglasHIS.His_ConsultarHojasRegistradas(ml_IdEstablecimientoActual, ml_IdLote).RecordCount) = 0 Then
                    MsgBox "Ya se ingreso el número maximo de hojas para este lote, por favor dígite otro lote", vbInformation, "HIS"
                    txtFechaAnio.Text = "____"
                    txtNroPaginas.Text = ""
                    txtUltimaPaginaLoteActiva.Text = ""
                    Exit Sub
                End If
                'txtUltimaPaginaLoteActiva.Text = mo_ReglasHIS.ObtenerDatosLoteNroHojaLibre(mo_DetalleLotes.IdLote)
                'Me.txtPagRestante.Text = Val(txtNroPaginas.Text) - Val(txtUltimaPaginaLoteActiva.Text)
                mo_Formulario.HabilitarDeshabilitar txtLote, False
                mb_SeleccionoLote = True
            Else
                mb_SeleccionoLote = False
            End If
            Set mo_DetalleLotes = Nothing
            Me.txtUltimaPaginaLoteActiva.SetFocus
        Case vbKeyReturn
            Dim oRcs_DetalleLotes As New Recordset
             If txtLote.Text <> "" Then
                  Set oRcs_DetalleLotes = mr_ReglasHIS.ConsultarRegistroFiltroLotes(ml_IdEstablecimientoActual, 0, 0, txtLote.Text, False)
                  If oRcs_DetalleLotes.RecordCount > 0 Then
                      oRcs_DetalleLotes.MoveFirst
                      ml_IdLote = oRcs_DetalleLotes.Fields!IdHisLote
                      txtFechaAnio.Text = oRcs_DetalleLotes.Fields!Anio
                      mo_cmbMes.BoundText = oRcs_DetalleLotes.Fields!IdMes
                      txtLote.Text = oRcs_DetalleLotes.Fields!Lote
                      txtNroPaginas.Text = oRcs_DetalleLotes.Fields!NroHojas
                      If mi_Opcion = sghAgregar Then
                        If Val(oRcs_DetalleLotes.Fields!NroHojas) - Val(mr_ReglasHIS.His_ConsultarHojasRegistradas(ml_IdEstablecimientoActual, ml_IdLote).RecordCount) = 0 Then
                            MsgBox "Ya se ingreso el número máximo de hojas para este lote, por favor dígite otro lote", vbInformation, Me.Caption
                            txtFechaAnio.Text = "____"
                            txtNroPaginas.Text = ""
                            txtUltimaPaginaLoteActiva.Text = ""
                            Exit Sub
                        End If
                      End If
                      mo_Formulario.HabilitarDeshabilitar txtLote, False
                      mb_SeleccionoLote = True
                      Me.txtUltimaPaginaLoteActiva.SetFocus
                  Else
                      mb_SeleccionoLote = False
                      Call MsgBox("El Lote ingresado no es válido, por favor oprima F11.", vbExclamation, Me.Caption)
                      txtLote.Text = ""
                  End If
              Else
                  mb_SeleccionoLote = False
                  mo_Teclado.RealizarNavegacion KeyCode, Me.txtUltimaPaginaLoteActiva
              End If
            Case vbKeyEscape
                If MsgBox("Desea salir del registro HIS?", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
                    btnCancelar_Click
                End If
    End Select
End Sub

Private Sub txtNacionalidad_LostFocus()
    txtNacionalidad.Text = UCase(txtNacionalidad.Text)
End Sub

Private Sub txtNroDocumento_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48) Or KeyAscii > 57 Then
        If KeyAscii = 8 Then
            KeyAscii = 8
        Else
            KeyAscii = 1
        End If
    End If
End Sub



Private Sub txtNroRegistro_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF11
        If Trim(Me.txtLote.Text) = "" Then
            MsgBox "No se digitó el lote", vbInformation, Me.Caption
            Exit Sub
        End If
        If Trim(Me.txtUltimaPaginaLoteActiva.Text) = "" Then
            MsgBox "No se digitó la hoja actual", vbInformation, Me.Caption
            Exit Sub
        End If
        If Trim(Me.txtResponsable.Text) = "" Then
            MsgBox "No se digitó el responsable", vbInformation, Me.Caption
            Exit Sub
        End If
        Dim mo_DetalleRegistros As New frmDetalleNroRegistrosLibres
        mo_DetalleRegistros.IdHisCabecera = ml_IdCabeceraHIS
        mo_DetalleRegistros.MostrarFormulario
        If mo_DetalleRegistros.BotonPresionado = sghAceptar Then
            Me.txtNroRegistro.Text = mo_DetalleRegistros.NroRegistros
            mo_Teclado.RealizarNavegacion KeyCode, txtDia
        Else
             Me.txtNroRegistro.SetFocus
        End If
        Set mo_DetalleRegistros = Nothing
    Case vbKeyReturn
        Me.txtDia.SetFocus
'        mo_Teclado.RealizarNavegacion KeyCode, txtDia
    Case vbKeyF12   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
        GrabaAtencionDiagnosticos
    Case vbKeyF7    'CANCELA EDICION DE ATENCION
        CancelaEdicionAtencion
    Case vbKeyEscape
        If MsgBox("Desea salir del registro de la Hoja HIS?", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
            btnCancelar_Click
        End If
    End Select
End Sub

Private Sub txtNroRegistro_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48) Or KeyAscii > 57 Then
        If KeyAscii = 8 Then
            KeyAscii = 8
        Else
            KeyAscii = 1
        End If
    End If
End Sub

Private Sub txtOrdenFamiliar_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 49) Or KeyAscii > 57 Then
        If KeyAscii = 8 Then
            KeyAscii = 8
        Else
            KeyAscii = 1
        End If
    End If
End Sub

Private Sub txtPeso_KeyPress(KeyAscii As Integer)
    If ((KeyAscii < 48) Or KeyAscii > 57) Then
        If KeyAscii = 8 Then
            KeyAscii = 8
        Else
            If KeyAscii = 46 Then
                KeyAscii = 46
            Else
                KeyAscii = 1
            End If
        End If
    End If
End Sub

'Activar busqueda de responsable de atención correspondiente al establecimiento actual
Private Sub txtResponsable_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF11
        If mb_SeleccionoLote Then
            If mb_SeleccionoHoja Then
                Dim oForm As New BuscaMedicosHis
                Dim oTablaMedico As New DOMedico
                Dim oTablaEmpleado As New DOEmpleado
                oForm.IdEstablecimiento = ml_IdEstablecimientoActual
                oForm.NombMedico = Me.txtResponsable.Text
                oForm.Anio = Me.txtFechaAnio.Text
                oForm.Mes = mo_cmbMes.BoundText
                oForm.MostrarFormulario
                If oForm.BotonPresionado = sghAceptar Then
                    ml_IdMedicoResponsable = oForm.IdMedico
                    Me.txtResponsable.Text = oForm.NombMedico
                    mo_cmbServicioCodigo.BoundText = oForm.IdServicio
                    mo_cmbTurno.BoundText = oForm.IdTurno
                    Dim oRcs_DiasProgramados As New Recordset
                    'Listar las Programaciones con Parametro de Estabelciemitno para Discriminar la Lista y Devolver Boolean
                    Set oRcs_DiasProgramados = mo_ReglasHIS.ListarProgramacionMedicaPorMedicoYEstablecimiento(ml_IdEstablecimientoActual, ml_IdMedicoResponsable, mo_cmbMes.BoundText, CInt(Me.txtFechaAnio.Text))
                    If oRcs_DiasProgramados.RecordCount <> 0 Then
                        oRcs_DiasProgramados.MoveFirst
                        mo_Formulario.HabilitarDeshabilitar txtResponsable, False
                        mb_SeleccionoMedico = True
                        cmdIngresarDiagnosticos.Enabled = True
                        'CODIGO DE INGRESO DE CABECERA - YA QUE ESTE ES INGRESADO EN LA HORA QUE SE ELIGE AL RESPONSABLE
                        '==============================================================
                        'Cargar Datos de la cabecera a este objeto
                        With oCabeceraAtencion
                            .IdHisCabecera = 0
                            .IdHisLote = ml_IdLote
                            .IdEstablecimiento = ml_IdEstablecimientoActual
                            .IdServicio = mo_cmbServicioCodigo.BoundText
                            .IdMedico = ml_IdMedicoResponsable
                            .FechaCreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL
                            .IdUsuario = ml_IdUsuario
                            .NroFormato = mo_ReglasHIS.ObtenerDatosNroFormatoLibre(CInt(Me.txtFechaAnio.Text), ml_IdEstablecimientoActual)
                            .NroHojaHis = Val(Me.txtUltimaPaginaLoteActiva.Text)
                            .IdEstadoHis = 1
                            .IdTurno = Val(Me.cmbTurno.ItemData(Me.cmbTurno.ListIndex))
                        End With
                        ml_IdCabeceraHIS = mo_ReglasHIS.IngresarHojaHIS(oCabeceraAtencion)
                        SeleccionaNroRegistroLibre
                        Me.txtDia.SetFocus
                    End If
                Else
                    mb_SeleccionoMedico = False
                End If
            Else
                Call MsgBox("Ingrese previamente la hoja", vbExclamation, Me.Caption)
            End If
        Else
            Call MsgBox("Ingrese previamente el lote", vbExclamation, Me.Caption)
        End If
    Case vbKeyEscape
        If MsgBox("Desea salir del registro HIS?", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
            btnCancelar_Click
        End If
End Select

End Sub

Private Sub txtOrdenFamiliar_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn
         If Val(mo_cmbTipoDocumento.BoundText) = 8 Then
                If txtOrdenFamiliar.Text <> "" Then
                    Dim oRcs_Temp As New Recordset
                    Dim oRcs_temp2 As New Recordset
                    Set oRcs_Temp = mo_ReglasHIS.HISPacientesFiltraPorNroDocumentoYtipo(Me.txtNroDocumento.Text, Val(mo_cmbTipoDocumento.BoundText))
                    If oRcs_Temp.RecordCount <> 0 Then
                        oRcs_Temp.MoveFirst
                        Do While Not oRcs_Temp.EOF
                            If Not IsNull(oRcs_Temp!NroHijo) Then
                                If Val(Me.txtOrdenFamiliar.Text) = Val(oRcs_Temp!NroHijo) Then
                                    If ml_IdEstablecimientoActual = oRcs_Temp!IdEstablecimiento Then
                                        Me.txtNroHC_FF_COD.Text = oRcs_Temp!nrohc_ff
                                    End If
                                     'Nacionalidad
                                     IdCodigoNacionalidad = CStr(oRcs_Temp!idnacionalidad)
                                     Set oRcs_temp2 = mo_ReglasHIS.ObtenerDatosCodNacPorIdNac(IdCodigoNacionalidad)
                                     oRcs_temp2.MoveFirst
                                     txtNacionalidad.Text = CStr(oRcs_temp2!Codigo)
                                     
                                     'IDPACIENTE
                                     ml_IdPacienteGalenHos = IIf(IsNull(oRcs_Temp!IdPacienteGalenHos), 0, oRcs_Temp!IdPacienteGalenHos)
                                                            
                                    'ETNIA
                                     If Not IsNull(oRcs_Temp!IdEtnia) Then
                                         mo_cmbEtnia.BoundText = CStr(oRcs_Temp!IdEtnia)
                                     End If
                                     
                                     'SEXO
                                     If Not IsNull(oRcs_Temp!Sexo) Then
                                         mo_cmbSexo.BoundText = CStr(oRcs_Temp!Sexo)
                                     End If
                                End If
                            End If

                            oRcs_Temp.MoveNext
                        Loop
                    Else
                        MsgBox "No se encontró dato alguno", vbInformation, Me.Caption
                    End If
                End If
            End If
        mo_Teclado.RealizarNavegacion KeyCode, cmbFinanciador
    Case vbKeyF12   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
        GrabaAtencionDiagnosticos
    Case vbKeyF7    'CANCELA EDICION DE ATENCION
        CancelaEdicionAtencion
    Case vbKeyEscape
        If MsgBox("Desea salir del registro de la Hoja HIS?", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
            btnCancelar_Click
        End If
    End Select
End Sub

Private Sub cmdIngresarDiagnosticos_Click()
    Dim Mensaje As String
    Mensaje = ValidarValoresAtencion
    If Len(Mensaje) = 0 Then
        mb_FaltaGrabarAtencion = True
        AdicionDiagnostico
        DeshabilitaPesoTallaMayores5Anios
    Else
        Call MsgBox("Existen los siguientes problemas:" + vbCrLf + Mensaje, vbInformation, Me.Caption)
        SeleccionaNroRegistroLibre
'        Me.txtDia.SetFocus
    End If
End Sub

'Activar la Busqueda del Tipo de Actividad que se edita en la Atencion Actual
Sub HabilitarDeshabilitarPorNroHcFFCod()
    Dim orsTemp As Recordset
    If Len(Trim(Me.txtNroHC_FF_COD.Text)) <= 6 Then
        Set orsTemp = mr_ReglasHIS.ObtenerListaCodigosActividadesporCodigoyNombre(Trim(Me.txtNroHC_FF_COD.Text), "")
        If orsTemp.RecordCount > 0 Then
            orsTemp.MoveFirst
                Do While Not orsTemp.EOF
                    If Trim(UCase(orsTemp.Fields!CodigoActividad)) = Trim(UCase(Me.txtNroHC_FF_COD.Text)) Then
                        IdTipoActividad = orsTemp.Fields!IdTipoAtencion
                        Exit Do
                    End If
                    orsTemp.MoveNext
                Loop
        End If
    End If
    HabilitarCamposAtencionPorActividad (IdTipoActividad)
End Sub

Private Sub txtNroHC_FF_COD_KeyUp(KeyCode As Integer, Shift As Integer)
   If txtNroHC_FF_COD.Text = "" Then
        IdTipoActividad = CInt(sghHISTipoActividad.Atencion)
   End If
End Sub

Private Sub txtNroHC_FF_COD_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF11
        Dim mo_frmCodigoActividad As New frmCodigoActividad
        mo_frmCodigoActividad.IdTipoActividad = ObtenerTipoActividad()
        mo_frmCodigoActividad.MostrarFormulario
        If mo_frmCodigoActividad.BotonPresionado = sghAceptar Then
            If Len(mo_frmCodigoActividad.CodigoSeleccionado) <> 0 Then
                Me.txtNroHC_FF_COD.Text = Trim(mo_frmCodigoActividad.CodigoSeleccionado)
                IdCodigoActividad = mo_frmCodigoActividad.IdCodigoSeleccionado
                IdTipoActividad = mo_frmCodigoActividad.IdTipoActividad
                'DETERMINAR LA HABILITACION DE LOS CAMPOS SEGUN SEA EL CASO DE TIPO DE ATENCION
                HabilitarCamposAtencionPorActividad (IdTipoActividad)
            Else
                HabilitarCamposAtencionPorActividad (sghHISTipoActividad.Atencion)
            End If
            Select Case IdTipoActividad
                Case sghHISTipoActividad.Atencion
                    txtNacionalidad.SetFocus
                Case sghHISTipoActividad.ActividadPreventivaPromocional, sghHISTipoActividad.ActividadConAnimales
                    cmdIngresarDiagnosticos.SetFocus
                Case sghHISTipoActividad.ActividadMasiva
                    txtEdad.SetFocus
            End Select
        Else
            IdTipoActividad = CInt(sghHISTipoActividad.Atencion)
            HabilitarDeshabilitarPorNroHcFFCod
'            HabilitarCamposAtencionPorActividad (sghHISTipoActividad.Atencion)
        End If
        Set mo_frmCodigoActividad = Nothing
    Case vbKeyReturn
        IdTipoActividad = CInt(sghHISTipoActividad.Atencion)
        HabilitarDeshabilitarPorNroHcFFCod
        Select Case IdTipoActividad
            Case sghHISTipoActividad.Atencion
                txtNacionalidad.SetFocus
            Case sghHISTipoActividad.ActividadPreventivaPromocional, sghHISTipoActividad.ActividadConAnimales
                cmdIngresarDiagnosticos.SetFocus
            Case sghHISTipoActividad.ActividadMasiva
                txtEdad.SetFocus
        End Select
    Case vbKeyF12   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
        GrabaAtencionDiagnosticos
    Case vbKeyF7    'CANCELA EDICION DE ATENCION
        CancelaEdicionAtencion
    Case vbKeyEscape
        If MsgBox("Desea salir del registro de la Hoja HIS?", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
            btnCancelar_Click
        End If
    End Select
End Sub

Private Sub txtNroHC_FF_COD_LostFocus()
    IdTipoActividad = CInt(sghHISTipoActividad.Atencion)
    If mb_SeleccionoMedico And mi_Opcion = sghAgregar Then
        HabilitarDeshabilitarPorNroHcFFCod
    End If
End Sub

'Activa la busqueda una Nacionaliadad de un Paciente
Private Sub txtNacionalidad_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF11
            Dim mo_frmDetalleCodigoPais As New frmDetalleCodigoPais
            mo_frmDetalleCodigoPais.MostrarFormulario
            
            If mo_frmDetalleCodigoPais.BotonPresionado = sghAceptar Then
                If mo_frmDetalleCodigoPais.CodigoNac <> "" Then
                    ml_IdNacionalidadAtencion = mo_frmDetalleCodigoPais.IdPais
                    Me.txtNacionalidad.Text = mo_frmDetalleCodigoPais.CodigoNac
                    Me.cmbTipoDocumento.SetFocus
                End If
            End If
            Set mo_frmDetalleCodigoPais = Nothing
        Case vbKeyReturn
            mo_Teclado.RealizarNavegacion KeyCode, Me.cmbTipoDocumento
        Case vbKeyF12   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
            GrabaAtencionDiagnosticos
        Case vbKeyF7    'CANCELA EDICION DE ATENCION
            CancelaEdicionAtencion
        Case vbKeyEscape
            If MsgBox("Desea salir del registro de la Hoja HIS?", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
                btnCancelar_Click
            End If
    End Select
End Sub

Private Sub txtDistritoProcedencia_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF11
            Dim oBusquedaDistrito As New SIGHNegocios.BuscarDistrito
            oBusquedaDistrito.IdDepartamentoBusqueda = ml_IdDepartamentoActual
            oBusquedaDistrito.IdProvinciaBusqueda = ml_IdProvinciaActual
            If Me.txtDistritoProcedencia.Text = "" Then
                oBusquedaDistrito.DescripcionDistrito = ""
            Else
                oBusquedaDistrito.DescripcionDistrito = Mid(Me.txtDistritoProcedencia.Text, 10, Len(Me.txtDistritoProcedencia.Text) - 9)
            End If
            oBusquedaDistrito.MostrarFormulario
            If oBusquedaDistrito.BotonPresionado = sghAceptar Then
                If oBusquedaDistrito.IdRegistroSeleccionado <> 0 Then
                    ml_IdDistritoAtencion = oBusquedaDistrito.IdRegistroSeleccionado
                    Me.txtDistritoProcedencia.Text = oBusquedaDistrito.IdRegistroSeleccionado & " - " & oBusquedaDistrito.DescripcionRegistroSeleccionado
                    If txtEdad.Enabled = True Then
                        Me.txtEdad.SetFocus
                    End If
                Else
                    ml_IdDistritoAtencion = 0
                    Me.txtDistritoProcedencia.Text = "No se eligio"
                End If
            End If
            Set oBusquedaDistrito = Nothing
        Case vbKeyBack
            ml_IdDistritoAtencion = 0
        Case vbKeyReturn
            If Me.txtEdad.Enabled = True Then Me.txtEdad.SetFocus
        Case vbKeyF12   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
            GrabaAtencionDiagnosticos
        Case vbKeyF7    'CANCELA EDICION DE ATENCION
            CancelaEdicionAtencion
        Case vbKeyEscape
            If MsgBox("Desea salir del registro de la Hoja HIS?", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
                btnCancelar_Click
            End If
    End Select
End Sub

Private Sub txtNroDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            If Me.txtNroDocumento.Text = "" Then
                Me.cmbFinanciador.SetFocus
                Exit Sub
            End If
            'Solo para DNI
            If Val(mo_cmbTipoDocumento.BoundText) = 1 Then
                Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
                Dim mo_HisGalenhos As New SIGHNegocios.ReglasHISGalenos
                Dim mo_DatosFechas As New SIGHEntidades.FechaHora
                Dim o_RcsDatosPaciente As New Recordset
        '        Dim o_RcsDatosPacienteSISAfiliacion As New Recordset
                
                Set o_RcsDatosPaciente = mo_AdminAdmision.PacientesFiltraPorNroDocumentoYtipo(Me.txtNroDocumento.Text, Val(mo_cmbTipoDocumento.BoundText))
        '        Set o_RcsDatosPacienteSISAfiliacion = mo_HisGalenhos.SisafilicianesObtenerNombreApellidosPorNroDocumentoYtipo(Me.txtNroDocumento.Text, Val(mo_cmbTipoDocumento.BoundText))
                
                'Si Encuentra apellidos y nombres del paciente en SISAfiliacion (SIGHExterna)
        '        If o_RcsDatosPacienteSISAfiliacion.RecordCount <> 0 Then
        '            o_RcsDatosPacienteSISAfiliacion.MoveFirst
        '        End If
                            
                'Si Encuentra algun Dato del Paciente
                If o_RcsDatosPaciente.RecordCount <> 0 Then
                    o_RcsDatosPaciente.MoveFirst
                    
                    'IDPACIENTE
                    ml_IdPacienteGalenHos = CLng(o_RcsDatosPaciente!IdPaciente)
                    
                    'ETNIA
                    If Not IsNull(o_RcsDatosPaciente!IdEtnia) Then
                        mo_cmbEtnia.BoundText = CStr(o_RcsDatosPaciente!IdEtnia)
                    End If
                    
                    'Distrito procedencia
                    If Not IsNull(o_RcsDatosPaciente!IdDistritoProcedencia) Then
                        ml_IdDistritoAtencion = CLng(o_RcsDatosPaciente!IdDistritoProcedencia)
                        
                        'busqueda de descripcion de distrito de procedencia
                        Dim mo_distrito As DODistrito
                        Set mo_distrito = mo_ReglasHIS.BusquedaDistrito(ml_IdDistritoAtencion)
                        Me.txtDistritoProcedencia = mo_distrito.Nombre
                    End If
                    
                    'Calculo del tipo de edad
                    If Not IsNull(o_RcsDatosPaciente!FechaNacimiento) Then
                        Dim md_fechanac As Date
                        Dim md_fechaActual As Date
                        Dim TipoEdad As Integer
                        Dim cantidadedad As Integer
                        
                        md_fechanac = CDate(o_RcsDatosPaciente!FechaNacimiento)
                        md_fechaActual = CDate(lcBuscaParametro.RetornaFechaServidorSQL)
                        
                        'verificamos si es un recien nacido
                        If mo_DatosFechas.CalculaSiEsRecienNacido(md_fechanac, md_fechaActual) = 1 Then
                            TipoEdad = CInt(sghHISTipoEdades.Dias) 'VERIFICAR TIPO DE EDAD EN DIAS
                            cantidadedad = mo_DatosFechas.EdadActualEnDias(md_fechanac, md_fechaActual)
                            
                        ElseIf mo_DatosFechas.DevuelveEdadEnMeses(md_fechanac, md_fechaActual) <= 11 Then
                            TipoEdad = CInt(sghHISTipoEdades.Meses) 'VERIFICAR TIPO DE EDAD EN MESES
                            cantidadedad = mo_DatosFechas.DevuelveEdadEnMeses(md_fechanac, md_fechaActual)
                        Else
                            TipoEdad = CInt(sghHISTipoEdades.Años)  'VERIFICAR TIPO DE EDAD EN AÑOS
                            cantidadedad = mo_DatosFechas.EdadActual(md_fechanac, md_fechaActual)
                        End If
                        
                        mo_cmbTipoEdad.BoundText = CStr(TipoEdad)
                        Me.txtEdad = Format(CStr(cantidadedad), "00")
                    End If
                    mo_cmbSexo.BoundText = CStr(o_RcsDatosPaciente!IdTipoSexo)
                Else
                    Dim oRcs_Temp As New Recordset
                    Dim oRcs_temp2 As New Recordset
                    
                    Set oRcs_Temp = mo_ReglasHIS.HISPacientesFiltraPorNroDocumentoYtipo(Me.txtNroDocumento.Text, Val(mo_cmbTipoDocumento.BoundText))
                    If oRcs_Temp.RecordCount <> 0 Then
                        oRcs_Temp.MoveFirst
                        'nrohc_ff
                        Do While Not oRcs_Temp.EOF
                            If ml_IdEstablecimientoActual = oRcs_Temp!IdEstablecimiento Then
                                Me.txtNroHC_FF_COD.Text = oRcs_Temp!nrohc_ff
                                Exit Do
                            End If
                            oRcs_Temp.MoveNext
                        Loop
                        oRcs_Temp.MoveFirst
                
                        'Nacionalidad
                        IdCodigoNacionalidad = CStr(oRcs_Temp!idnacionalidad)
                        Set oRcs_temp2 = mo_ReglasHIS.ObtenerDatosCodNacPorIdNac(IdCodigoNacionalidad)
                        oRcs_temp2.MoveFirst
                        txtNacionalidad.Text = CStr(oRcs_temp2!Codigo)
                        
                        'IDPACIENTE
                        ml_IdPacienteGalenHos = IIf(IsNull(oRcs_Temp!IdPacienteGalenHos), 0, oRcs_Temp!IdPacienteGalenHos)
                       
                        'Nro Hijo
                        If Val(mo_cmbTipoDocumento.BoundText) = 8 Then
                            If Not IsNull(oRcs_Temp!NroHijo) Then
                                Me.txtOrdenFamiliar.Text = CStr(oRcs_Temp!NroHijo)
                            End If
                        End If
                        
                       'ETNIA
                        If Not IsNull(oRcs_Temp!IdEtnia) Then
                            mo_cmbEtnia.BoundText = CStr(oRcs_Temp!IdEtnia)
                        End If
                        
                        'SEXO
                        If Not IsNull(oRcs_Temp!Sexo) Then
                            mo_cmbSexo.BoundText = CStr(oRcs_Temp!Sexo)
                        End If
                        cmbFinanciador.SetFocus
                    Else
                        MsgBox "No se encontró dato alguno", vbInformation, Me.Caption
                    End If
                End If
                txtOrdenFamiliar.SetFocus
            Else
                If Val(mo_cmbTipoDocumento.BoundText) = 8 Then
                    Me.txtOrdenFamiliar.SetFocus
                Else
                    Me.cmbFinanciador.SetFocus
                End If
            End If
        Case vbKeyF12   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
            GrabaAtencionDiagnosticos
        Case vbKeyF7    'CANCELA EDICION DE ATENCION
            CancelaEdicionAtencion
        Case vbKeyEscape
            If MsgBox("Desea salir del registro de la Hoja HIS?", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
                btnCancelar_Click
            End If
    End Select
End Sub

Private Sub txtNroDocumento_GotFocus()
    Me.lblMensajeDNIBusqueda.Caption = "INGRESE EL DNI Y PRESIONE ENTER PARA BUSCAR DATOS"
End Sub

Private Sub txtNroDocumento_LostFocus()
    Me.lblMensajeDNIBusqueda.Caption = ""
End Sub

'------------------------- EVENTOS PARA LOS LISTADOS DEL FORMULARIO ------------------------
Private Sub cmbTipoDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            mo_Teclado.RealizarNavegacion KeyCode, cmbTipoDocumento
        Case vbKeyF12   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
            GrabaAtencionDiagnosticos
        Case vbKeyF7    'CANCELA EDICION DE ATENCION
            CancelaEdicionAtencion
        Case vbKeyEscape
            If MsgBox("Desea salir del registro de la Hoja HIS?", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
                btnCancelar_Click
            End If
    End Select
End Sub

Private Sub cmbTipoDocumento_LostFocus()
   If cmbTipoDocumento.Text <> "" Then
        On Error Resume Next
        mo_cmbTipoDocumento.BoundText = Val(Split(cmbTipoDocumento.Text, " = ")(0))
        If mo_cmbTipoDocumento.BoundText = 8 Then
           mo_Formulario.HabilitarDeshabilitar Me.txtOrdenFamiliar, True
        Else
            Me.txtOrdenFamiliar.Text = ""
            mo_Formulario.HabilitarDeshabilitar Me.txtOrdenFamiliar, False
        End If
   End If

End Sub

Private Sub cmbTipoDocumento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

'cmbFinanciador
Private Sub cmbFinanciador_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            If ml_IdPacienteGalenHos = 0 Then
                cmbEtnia.SetFocus
            Else
                cmbEtnia.SetFocus
            End If
        Case vbKeyF12   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
            GrabaAtencionDiagnosticos
        Case vbKeyF7    'CANCELA EDICION DE ATENCION
            CancelaEdicionAtencion
        Case vbKeyEscape
            If MsgBox("Desea salir del registro de la Hoja HIS?", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
                btnCancelar_Click
            End If
    End Select
End Sub

Private Sub cmbFinanciador_LostFocus()
   If cmbFinanciador.Text <> "" Then
        On Error Resume Next
       mo_cmbFinanciador.BoundText = Val(Split(cmbFinanciador.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbFinanciador
End Sub

Private Sub cmbFinanciador_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

'cmbEtnia
Private Sub cmbEtnia_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            mo_Teclado.RealizarNavegacion KeyCode, Me.txtDistritoProcedencia
        Case vbKeyF12   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
            GrabaAtencionDiagnosticos
        Case vbKeyF7    'CANCELA EDICION DE ATENCION
            CancelaEdicionAtencion
        Case vbKeyEscape
            If MsgBox("Desea salir del registro de la Hoja HIS?", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
                btnCancelar_Click
            End If
    End Select
End Sub

Private Sub cmbEtnia_LostFocus()
   If cmbEtnia.Text <> "" Then
        On Error Resume Next
       mo_cmbEtnia.BoundText = Val(Split(cmbEtnia.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbEtnia
End Sub

Private Sub cmbEtnia_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub cmbTipoEdad_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF12   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
            GrabaAtencionDiagnosticos
        Case vbKeyF7    'CANCELA EDICION DE ATENCION
            CancelaEdicionAtencion
        Case vbKeyEscape
            If MsgBox("Desea salir del registro de la Hoja HIS?", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
                btnCancelar_Click
            End If
        Case vbKeyReturn
            DeshabilitaPesoTallaMayores5Anios
            mo_Teclado.RealizarNavegacion KeyCode, Me.cmbSexo
        Case Else
            DeshabilitaPesoTallaMayores5Anios
            mo_Teclado.RealizarNavegacion KeyCode, Me.cmbSexo
    End Select
End Sub

Private Sub DeshabilitaPesoTallaMayores5Anios()
    If IdTipoActividad = sghHISTipoActividad.Atencion Then
        If lcBuscaParametro.SeleccionaFilaParametro(329) = "S" Then
            If CInt(IIf(Replace(Me.txtEdad.Text, "_", "") = "", 0, Replace(Me.txtEdad.Text, "_", ""))) <> 0 Then
                If CInt(Val(mo_cmbTipoEdad.BoundText)) = sghHISTipoEdades.Dias Or CInt(Val(mo_cmbTipoEdad.BoundText)) = sghHISTipoEdades.Meses Then
                     mo_Formulario.HabilitarDeshabilitar Me.txtPeso, True
                     mo_Formulario.HabilitarDeshabilitar Me.txtTalla, True
                     mb_PesoTallaHabilitados = True
                Else
                    If CInt(IIf(Replace(Me.txtEdad.Text, "_", "") = "", 0, Replace(Me.txtEdad.Text, "_", ""))) < 5 Then
                        mo_Formulario.HabilitarDeshabilitar Me.txtPeso, True
                        mo_Formulario.HabilitarDeshabilitar Me.txtTalla, True
                        mb_PesoTallaHabilitados = True
                    Else
                        mo_Formulario.HabilitarDeshabilitar Me.txtPeso, False
                        Me.txtPeso.Text = ""
                        mo_Formulario.HabilitarDeshabilitar Me.txtTalla, False
                        mb_PesoTallaHabilitados = False
                        Me.txtTalla.Text = ""
                    End If
                End If
            End If
        Else
            mo_Formulario.HabilitarDeshabilitar Me.txtPeso, False
            mo_Formulario.HabilitarDeshabilitar Me.txtTalla, False
        End If
    End If
End Sub
Private Sub txtEdad_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            mo_Teclado.RealizarNavegacion KeyCode, Me.cmbTipoEdad
        Case vbKeyF12   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
            GrabaAtencionDiagnosticos
        Case vbKeyF7    'CANCELA EDICION DE ATENCION
            CancelaEdicionAtencion
        Case vbKeyEscape
            If MsgBox("Desea salir del registro de la Hoja HIS?", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
                btnCancelar_Click
            End If
    End Select
End Sub

Private Sub txtPeso_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn
        mo_Teclado.RealizarNavegacion KeyCode, txtTalla
    Case vbKeyF12   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
        GrabaAtencionDiagnosticos
    Case vbKeyF7    'CANCELA EDICION DE ATENCION
        CancelaEdicionAtencion
    Case vbKeyEscape
        If MsgBox("Desea salir del registro de la Hoja HIS?", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
            btnCancelar_Click
        End If
    End Select
End Sub

Private Sub txtResponsable_LostFocus()
    If txtResponsable <> "" Then
        mo_Formulario.HabilitarDeshabilitar Me.txtNroRegistro, True
        mo_Formulario.HabilitarDeshabilitar Me.txtDia, True
        mo_Formulario.HabilitarDeshabilitar Me.txtNroHC_FF_COD, True
        mo_Formulario.HabilitarDeshabilitar Me.txtNacionalidad, True
        mo_Formulario.HabilitarDeshabilitar Me.cmbTipoDocumento, True
        mo_Formulario.HabilitarDeshabilitar Me.txtNroDocumento, True
        mo_Formulario.HabilitarDeshabilitar Me.txtOrdenFamiliar, False
        mo_Formulario.HabilitarDeshabilitar Me.cmbFinanciador, True
        mo_Formulario.HabilitarDeshabilitar Me.cmbEtnia, True
        mo_Formulario.HabilitarDeshabilitar Me.txtDistritoProcedencia, True
        mo_Formulario.HabilitarDeshabilitar Me.txtEdad, True
        mo_Formulario.HabilitarDeshabilitar Me.cmbTipoEdad, True
        mo_Formulario.HabilitarDeshabilitar Me.cmbSexo, True
        
        If lcBuscaParametro.SeleccionaFilaParametro(329) = "S" Then
            mo_Formulario.HabilitarDeshabilitar Me.txtPeso, True
            mo_Formulario.HabilitarDeshabilitar Me.txtTalla, True
        Else
            mo_Formulario.HabilitarDeshabilitar Me.txtPeso, False
            mo_Formulario.HabilitarDeshabilitar Me.txtTalla, False
        End If
        
        mo_Formulario.HabilitarDeshabilitar Me.cmbEstadoFrenteEstablecimiento, True
        mo_Formulario.HabilitarDeshabilitar Me.cmbEstadoFrenteServicio, True
    End If
End Sub

Private Sub txtTalla_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn
        mo_Teclado.RealizarNavegacion KeyCode, cmbEstadoFrenteServicio
    Case vbKeyF12   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
        GrabaAtencionDiagnosticos
    Case vbKeyF7    'CANCELA EDICION DE ATENCION
        CancelaEdicionAtencion
    Case vbKeyEscape
        If MsgBox("Desea salir del registro de la Hoja HIS?", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
            btnCancelar_Click
        End If
    End Select
End Sub

'cmbSexo
Private Sub cmbSexo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            mo_Teclado.RealizarNavegacion KeyCode, txtPeso
        Case vbKeyF12   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
            GrabaAtencionDiagnosticos
        Case vbKeyF7    'CANCELA EDICION DE ATENCION
            CancelaEdicionAtencion
        Case vbKeyEscape
            If MsgBox("Desea salir del registro de la Hoja HIS?", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
                btnCancelar_Click
            End If
    End Select
End Sub

Private Sub cmbSexo_LostFocus()
   If cmbSexo.Text = "" Then
       mo_Formulario.MarcarComoVacio cmbSexo
   End If
End Sub

'cmbEstadoFrenteEstablecimiento
Private Sub cmbEstadoFrenteEstablecimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn
        mo_Teclado.RealizarNavegacion KeyCode, cmdIngresarDiagnosticos
    Case vbKeyF12   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
        GrabaAtencionDiagnosticos
    Case vbKeyF7    'CANCELA EDICION DE ATENCION
        CancelaEdicionAtencion
    Case vbKeyEscape
        If MsgBox("Desea salir del registro de la Hoja HIS?", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
            btnCancelar_Click
        End If
    End Select
End Sub

Private Sub cmbEstadoFrenteEstablecimiento_LostFocus()
   If cmbEstadoFrenteEstablecimiento.Text = "" Then
      mo_Formulario.MarcarComoVacio cmbEstadoFrenteEstablecimiento
   End If
End Sub

Private Sub cmbEstadoFrenteServicio_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn
        mo_Teclado.RealizarNavegacion KeyCode, cmbEstadoFrenteEstablecimiento
    Case vbKeyF12   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
        GrabaAtencionDiagnosticos
    Case vbKeyF7    'CANCELA EDICION DE ATENCION
        CancelaEdicionAtencion
    Case vbKeyEscape
        If MsgBox("Desea salir del registro de la Hoja HIS?", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
            btnCancelar_Click
        End If
    End Select
End Sub

Private Sub cmbEstadoFrenteServicio_LostFocus()
   If cmbEstadoFrenteServicio.Text = "" Then
       mo_Formulario.MarcarComoVacio cmbEstadoFrenteServicio
   End If
End Sub

Private Sub txtUltimaPaginaLoteActiva_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF11
            If mb_SeleccionoLote Then
                Dim mo_DetalleHojasLibres As New frmDetalleHojasLibres
                mo_DetalleHojasLibres.IdEstablecimiento = ml_IdEstablecimientoActual
                mo_DetalleHojasLibres.IdLote = ml_IdLote
                mo_DetalleHojasLibres.TotalPaginas = Val(txtNroPaginas.Text)
                mo_DetalleHojasLibres.MostrarFormulario
                If mo_DetalleHojasLibres.BotonPresionado = sghAceptar Then
                    txtUltimaPaginaLoteActiva.Text = mo_DetalleHojasLibres.NumeroHoja
                    mo_Formulario.HabilitarDeshabilitar Me.txtUltimaPaginaLoteActiva, False
                    mb_SeleccionoHoja = True
                    Set mo_DetalleHojasLibres = Nothing
                    Me.txtResponsable.SetFocus
                Else
                    Set mo_DetalleHojasLibres = Nothing
                    Me.txtUltimaPaginaLoteActiva.SetFocus
                End If

            Else
                mb_SeleccionoHoja = False
                Call MsgBox("Ingrese previamente el lote", vbExclamation, Me.Caption)
            End If
        Case vbKeyReturn
            If mb_SeleccionoLote Then
                If mi_Opcion = sghAgregar Then
                    Dim HojaUsada As Boolean
                    Dim oRcsTemp As New Recordset
                    Set oRcsTemp = mr_ReglasHIS.His_ConsultarHojasRegistradas(ml_IdEstablecimientoActual, ml_IdLote)
                    HojaUsada = False
                    If oRcsTemp.RecordCount > 0 Then
                        oRcsTemp.MoveFirst
                        Do While Not oRcsTemp.EOF
                            If Val(Me.txtUltimaPaginaLoteActiva.Text) = Int(oRcsTemp!NroHojaHis) Then
                                HojaUsada = True
                            End If
                            oRcsTemp.MoveNext
                        Loop
                    End If
                    If HojaUsada = False Then
                        If Me.txtUltimaPaginaLoteActiva.Text <> "" Then
                            If Val(Me.txtUltimaPaginaLoteActiva.Text) > Val(Me.txtNroPaginas.Text) Then
                                Call MsgBox("La hoja actual no puede ser mayor al total de hojas", vbExclamation, Me.Caption)
                            Else
                                mb_SeleccionoHoja = True
                                mo_Formulario.HabilitarDeshabilitar Me.txtUltimaPaginaLoteActiva, False
                                Me.txtResponsable.SetFocus
                            End If
                        Else
                            Me.txtUltimaPaginaLoteActiva.SetFocus
                        End If
                    Else
                        mb_SeleccionoHoja = False
                        Call MsgBox("La hoja ya fue registrada, por favor oprima F11.", vbExclamation, Me.Caption)
                    End If
                End If
            Else
                Call MsgBox("Ingrese previamente el lote", vbExclamation, Me.Caption)
                Me.txtLote.SetFocus
            End If
        Case vbKeyEscape
            If MsgBox("Desea salir del registro HIS?", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
                btnCancelar_Click
            End If
    End Select
End Sub

Private Sub txtUltimaPaginaLoteActiva_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48) Or KeyAscii > 57 Then
        If KeyAscii = 8 Then
            KeyAscii = 8
        Else
            KeyAscii = 1
        End If
    End If
End Sub

Private Sub ugvDetalleDiagnosticos_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If Not ugvDetalleDiagnosticos.ActiveCell Is Nothing Then
        If ugvDetalleDiagnosticos.ActiveCell.Column.Index = 5 Then
            If ugvDetalleDiagnosticos.ActiveCell.GetText <> "" Then
                If KeyAscii = 8 Then
                    KeyAscii = 8
                Else
                    If Len(CStr(ugvDetalleDiagnosticos.ActiveCell.GetText)) > 3 Then
                        KeyAscii = 1
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub txtTalla_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48) Or KeyAscii > 57 Then
        If KeyAscii = 8 Then
            KeyAscii = 8
        Else
            KeyAscii = 1
        End If
    End If
End Sub

Private Sub ugvDetalleDiagnosticos_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    If mi_Opcion = sghConsultar Or mi_Opcion = sghEliminar Then
        Exit Sub
    End If
    Select Case KeyCode
    Case vbKeyEscape
        Me.ugvDetalleDiagnosticos.Update
        If MsgBox("Desea salir del registro HIS?", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
            btnCancelar_Click
        End If
    Case vbKeyF10   'ADICIONA UN DIAGNOSTICO NUEVO
        Me.ugvDetalleDiagnosticos.Update
        AdicionDiagnostico
    Case vbKeyF11   'CONSULTA LOS LISTADOS CORRESPONDIENTES
        ConsultaListadoCorrespondiente
    Case vbKeyF12   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
        GrabaAtencionDiagnosticos
    Case vbKeyF7    'CANCELA EDICION DE ATENCION
        CancelaEdicionAtencion
    Case vbKeyF6    'BORRA EL DIAGNOSTICO
        If oRcs_Diagnosticos.RecordCount > 0 Then
            If MsgBox("Desea eliminar el diagnóstico actual?", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
                EliminarDiagnostico
            End If
        End If
    Case vbKeyReturn
        SaltarConEnterColumnaDiagnosticos
    End Select
End Sub

Public Sub SaltarConEnterColumnaDiagnosticos()
    If oRcs_Diagnosticos.RecordCount <> 0 Then
        If Not Me.ugvDetalleDiagnosticos.ActiveCell Is Nothing Then
            If Not IsNull(Me.ugvDetalleDiagnosticos.ActiveCell.Column.DataField) Then
                Select Case Me.ugvDetalleDiagnosticos.ActiveCell.Column.DataField
                Case "DESCRIPCION_CIE", "IdSubClasificacionDX"
                    SendKeys "{tab}"
                Case "CodLAB"
                    cmdIngresarDiagnosticos.SetFocus
                End Select
            End If
        End If
    End If
End Sub

Public Sub ConsultaListadoCorrespondiente()
        If oRcs_Diagnosticos.RecordCount <> 0 Then
           If Not Me.ugvDetalleDiagnosticos.ActiveCell Is Nothing Then
            If Not IsNull(Me.ugvDetalleDiagnosticos.ActiveCell.Column.DataField) Then
                'PARA EL LISTADO DE CODIGOS LAB
                If Me.ugvDetalleDiagnosticos.ActiveCell.Column.DataField = "CodLAB" Then
                    Dim oForm As New frmDetalleCodigosLAB
                    oForm.MostrarFormulario
                    If oForm.BotonPresionado = sghAceptar Then
                        If Trim(oForm.CodigoLab) <> "" Then 'Actualizado 01102014
                            Me.ugvDetalleDiagnosticos.ActiveRow.Cells("CodLAB").Value = oForm.CodigoLab
                        End If
                    End If
                End If
                'PARA EL LISTADO DE CODIGOS CIE 10
                If Me.ugvDetalleDiagnosticos.ActiveCell.Column.DataField = "DESCRIPCION_CIE" Then
                    Me.ugvDetalleDiagnosticos.ActiveRow.Cells("IdCIE").Value = 0
                    Dim oBusqueda As New SIGHhisDigitacion.BusquedaProductosHis
                    Dim oDoFACTCATALOGOSERVICIOS As New sighcomun.DOHis_FactCatalogoServicios
                    Me.ugvDetalleDiagnosticos.Update
                    If IsNull(Me.ugvDetalleDiagnosticos.ActiveRow.Cells("DESCRIPCION_CIE").Value) Then
                        oBusqueda.CodigoDx = ""
                    Else
                        oBusqueda.CodigoDx = CStr(Me.ugvDetalleDiagnosticos.ActiveRow.Cells("DESCRIPCION_CIE").Value)
                    End If
                    oBusqueda.MostrarFormulario
                    
                    If oBusqueda.BotonPresionado = sghAceptar Then
                        Dim mo_AdminServiciosComunes As New SIGHDatos.His_FactCatalogoServicios
                        oDoFACTCATALOGOSERVICIOS.IdDiagCpt = oBusqueda.IdRegistroSeleccionado
                        
                        If mo_AdminServiciosComunes.SeleccionarPorId(oDoFACTCATALOGOSERVICIOS) Then
                            'VALIDACION DE DIAGNOSTICOS REPETIDOS
                            If oRcs_Diagnosticos.RecordCount <> 0 Then
                            'SE CLONARA PARA BUSCAR EL DATO
                                Dim oRcs_Temp As New Recordset
                                Dim oRcs_temp2 As New Recordset
                                
                                Set oRcs_Temp = oRcs_Diagnosticos.Clone(adLockReadOnly)
                                oRcs_Temp.Filter = "IdHisDetalle=" & IdAtencion
                                oRcs_Temp.MoveFirst
                                
                                Do While Not oRcs_Temp.EOF
                                    If Not IsNull(oRcs_Temp!IdCIE) Then 'indica si es el primero ingresado
                                        If oDoFACTCATALOGOSERVICIOS.IdDiagCpt = CLng(oRcs_Temp!IdCIE) Then
                                           Set oRcs_temp2 = mo_ReglasHIS.DevuelveCodigoDiagnosticosHis(oDoFACTCATALOGOSERVICIOS.IdDiagCpt)
                                                If oRcs_temp2.RecordCount > 0 Then
                                                    If oRcs_temp2!MasDeUnDiagnosticos = 0 Then
                                                        Call MsgBox("El Producto His ya fue ingresado.", vbExclamation Or vbSystemModal, Me.Caption)
                                                         Exit Sub
                                                    End If
                                                End If
                                        End If
                                    End If
                                    oRcs_Temp.MoveNext
                                Loop
                                oRcs_Temp.Close
                                Set oRcs_Temp = Nothing
                            End If
                            
                            'VALIDACION DE DIAGNOSTICOS DEPENDIENDO DE LOS DATOS DEL PACIENTE (GENERO-EDAD)
                            Dim ms_MensajeErrorDiagnostico As String
'                            ms_MensajeErrorDiagnostico = ValidacionDiagnostico(oDODiagnostico)
                            ms_MensajeErrorDiagnostico = ""
                            
                            If Len(ms_MensajeErrorDiagnostico) = 0 Then
                                Me.ugvDetalleDiagnosticos.ActiveRow.Cells("IdCIE").Value = CLng(oDoFACTCATALOGOSERVICIOS.IdDiagCpt)
                                Me.ugvDetalleDiagnosticos.ActiveRow.Cells("DESCRIPCION_CIE").Value = "(" & Trim(CStr(oDoFACTCATALOGOSERVICIOS.codigodiagcpt)) & ") - " & CStr(oDoFACTCATALOGOSERVICIOS.descripciondiagcpt)
                            Else
                                Call MsgBox("Se verifico lo siguiente " & vbCrLf & ms_MensajeErrorDiagnostico, vbExclamation Or vbSystemModal, Me.Caption)
                                Exit Sub
                            End If
                        Else
                            Me.ugvDetalleDiagnosticos.ActiveRow.Cells("IdCIE").Value = ""
                            Me.ugvDetalleDiagnosticos.ActiveRow.Cells("DESCRIPCION_CIE").Value = "FALTA DIAGNOSTICO"
                            Me.ugvDetalleDiagnosticos.ActiveRow.Cells("DESCRIPCION_CIE").Activation = ssActivationAllowEdit
                        End If
                    End If
                    Set oBusqueda = Nothing
                End If
            End If
          End If
        End If
End Sub

Private Sub ugvResumenHIS_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
Select Case KeyCode
Case vbKeyF8    'MODIFICACION DE REGISTRO DE ATENCION
    If MsgBox("¿Desea modificar esta atención?.", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
        CargarAtencion
        HabilitarCamposAtencionPorActividad IdTipoActividad
        DeshabilitaPesoTallaMayores5Anios
        mb_FaltaGrabarAtencion = True
        Me.txtDia.SetFocus
    End If
Case vbKeyF6    'ELIMINACION DE REGISTRO DE ATENCION
    If MsgBox("Desea eliminar la atención con todos sus diagnósticos?", vbYesNo Or vbQuestion Or vbSystemModal Or vbDefaultButton1, Me.Caption) = vbYes Then
        If EliminarAtencion Then
            MsgBox "Se eliminó la atención correctamente.", vbExclamation, Me.Caption
            RefrescarListaAtenciones
        Else
            MsgBox "No se eliminó la atención.", vbCritical, Me.Caption
        End If
    End If
End Select
End Sub

Private Sub ugvResumenHIS_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
Layout.Override.RowSizingArea = ssRowSizingAreaEntireRow
Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
Layout.Override.AllowDelete = ssAllowDeleteNo

With Me.ugvResumenHIS.Bands(0)
    'Configuracion de Detalle de Atenciones
    .Columns("IdHisCabecera").Hidden = True
    .Columns("IdHisDetalle").Hidden = True
    .Columns("IdPacienteGalenHos").Hidden = True
    .Columns("IdTipoAtencion").Hidden = True
    
    .Columns("NroRegistroLote").Hidden = True
    .Columns("NroRegistroHoja").Header.Caption = "Nro Registro"
    .Columns("NroRegistroHoja").Width = 1000
    .Columns("NroRegistroHoja").Activation = ssActivationActivateNoEdit
       
    .Columns("DiaAtencion").Header.Caption = "Dia"
    .Columns("DiaAtencion").Width = 500
    .Columns("DiaAtencion").Activation = ssActivationActivateNoEdit
       
    .Columns("IdHisPaciente").Hidden = True
    .Columns("IdNacionalidad").Hidden = True
       
    .Columns("HC_FF_COD").Header.Caption = "HC / FF / Cod.Act"
    .Columns("HC_FF_COD").Width = 1800
    .Columns("HC_FF_COD").Activation = ssActivationActivateNoEdit
       
    .Columns("IdNacionalidad").Header.Caption = "Nacionalidad"
    .Columns("IdNacionalidad").Width = 1800
    .Columns("IdNacionalidad").Activation = ssActivationActivateNoEdit
    
    .Columns("IdTipoDocIdentidad").Header.Caption = "Tipo Doc."
    .Columns("IdTipoDocIdentidad").Width = 1000
    .Columns("IdTipoDocIdentidad").ValueList = "TiposDocIdentidad"
    .Columns("IdTipoDocIdentidad").Activation = ssActivationActivateNoEdit
    
    .Columns("NroDocIdentidad").Header.Caption = "Nro Doc."
    .Columns("NroDocIdentidad").Width = 1300
    .Columns("NroDocIdentidad").Activation = ssActivationActivateNoEdit
    
    .Columns("NroHijo").Header.Caption = "Nro Hijo"
    .Columns("NroHijo").Width = 1000
    .Columns("NroHijo").Activation = ssActivationActivateNoEdit
    
    .Columns("IdFinanciador").Header.Caption = "Tipo Financiamiento"
    .Columns("IdFinanciador").Width = 2000
    .Columns("IdFinanciador").ValueList = "FuentesFinanciamiento"
    .Columns("IdFinanciador").Activation = ssActivationActivateNoEdit

    .Columns("IdEtnia").Header.Caption = "Etnia"
    .Columns("IdEtnia").Width = 1500
    .Columns("IdEtnia").ValueList = "Etnias"
    .Columns("IdEtnia").Activation = ssActivationActivateNoEdit
    
    .Columns("IdDistrito").Hidden = True
    
    .Columns("TipoEdad").Header.Caption = "Tipo Edad"
    .Columns("TipoEdad").Width = 1000
    .Columns("TipoEdad").ValueList = "TiposEdad"
    .Columns("TipoEdad").Activation = ssActivationActivateNoEdit
    
    .Columns("Edad").Header.Caption = "Edad"
    .Columns("Edad").Width = 1000
    .Columns("Edad").Activation = ssActivationActivateNoEdit
    
    .Columns("Sexo").Header.Caption = "Sexo"
    .Columns("Sexo").Width = 1000
    .Columns("Sexo").ValueList = "Genero"
    .Columns("Sexo").Activation = ssActivationActivateNoEdit
        
    .Columns("Talla").Header.Caption = "Talla"
    .Columns("Talla").Width = 1000
    .Columns("Talla").Activation = ssActivationActivateNoEdit
    
    .Columns("Peso").Header.Caption = "Peso"
    .Columns("Peso").Width = 1000
    .Columns("Peso").Activation = ssActivationActivateNoEdit
        
    .Columns("IdEstadoaEstablec").Header.Caption = "Estado a Establec."
    .Columns("IdEstadoaEstablec").Width = 1500
    .Columns("IdEstadoaEstablec").ValueList = "EstadoFrenteEstablecimiento"
    .Columns("IdEstadoaEstablec").Activation = ssActivationActivateNoEdit
    
    .Columns("IdEstadoaServicio").Header.Caption = "Estado a Servicio"
    .Columns("IdEstadoaServicio").Width = 1500
    .Columns("IdEstadoaServicio").ValueList = "EstadoFrenteServicio"
    .Columns("IdEstadoaServicio").Activation = ssActivationActivateNoEdit
    
    .Columns("IdEstado").Hidden = True
End With
End Sub

Private Sub ugvResumenDiagnosticos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.RowSizingArea = ssRowSizingAreaEntireRow
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    Layout.Override.AllowDelete = ssAllowDeleteNo
    
    With Me.ugvResumenDiagnosticos.Bands(0)
        .Columns("IdHisDetalleDiagnostico").Hidden = True
        .Columns("IdHisDetalle").Hidden = True
    
        .Columns("IdCIE").Hidden = True
        .Columns("DESCRIPCION_CIE").Header.Caption = "Descripcion CIE"
        .Columns("DESCRIPCION_CIE").Width = 4000
    
        .Columns("IdSubClasificacionDX").Header.Caption = "Tipo Diagnostico"
        .Columns("IdSubClasificacionDX").Width = 2000
        .Columns("IdSubClasificacionDX").ValueList = "ClasificacionDiagnostico"
        .Columns("IdSubClasificacionDX").Style = ssStyleDropDownValidate
    
        .Columns("CodLAB").Header.Caption = "Codigo LAB"
        .Columns("CodLAB").Width = 2000
    
        .Columns("MSG_ALERTA").Header.Caption = "Mensaje Alerta"
        .Columns("MSG_ALERTA").Width = 2000
        .Columns("MSG_ALERTA").Activation = ssActivationActivateNoEdit
        .Columns("MSG_ALERTA").CellAppearance.BackColor = ml_ColorMensaje
        '.Columns("MSG_ALERTA").CellAppearance.ForeColor =
        .Columns("IdEstado").Hidden = True
    End With
End Sub

Private Sub ugvDetalleDiagnosticos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.RowSizingArea = ssRowSizingAreaEntireRow
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    Layout.Override.AllowDelete = ssAllowDeleteNo
    
    With Me.ugvDetalleDiagnosticos.Bands(0)
        .Columns("IdHisDetalleDiagnostico").Hidden = True
        .Columns("IdHisDetalle").Hidden = True
    
        .Columns("IdCIE").Hidden = True
        .Columns("DESCRIPCION_CIE").Header.Caption = "Descripcion CIE"
        .Columns("DESCRIPCION_CIE").Width = 4000
    
        .Columns("IdSubClasificacionDX").Header.Caption = "Tipo Diagnostico"
        .Columns("IdSubClasificacionDX").Width = 2000
        
        .Columns("IdSubClasificacionDX").ValueList = "ClasificacionDiagnostico"
        .Columns("IdSubClasificacionDX").Style = ssStyleDropDownValidate
    
        .Columns("CodLAB").Header.Caption = "Codigo LAB"
        .Columns("CodLAB").Width = 2000
    
        .Columns("MSG_ALERTA").Header.Caption = "Mensaje Alerta"
        .Columns("MSG_ALERTA").Width = 2000
        .Columns("MSG_ALERTA").Activation = ssActivationActivateNoEdit
        .Columns("MSG_ALERTA").CellAppearance.BackColor = ml_ColorMensaje
        '.Columns("MSG_ALERTA").CellAppearance.ForeColor =
        .Columns("IdEstado").Hidden = True
    End With
End Sub

'VISUALIZACION DE DIAGNOSTICOS POR CADA ATENCION
Private Sub ugvResumenHIS_AfterRowActivate()
    Dim Row As UltraGrid.SSRow
    Dim IdDetalleHIS As Long
    Set Row = Me.ugvResumenHIS.ActiveRow
    If Not IsNull(Row.Cells("IdHisDetalle").Value) Then
        IdDetalleHIS = CLng(Row.Cells("IdHisDetalle").Value)
        Set oRcs_DiagnosticosTemp.DataSource = mo_ReglasHIS.ObtenerDatosDetalleDiagnosticoPorIdDetalle(IdDetalleHIS)
        Set Me.ugvResumenDiagnosticos.DataSource = oRcs_DiagnosticosTemp
        'ugvResumenDiagnosticos.Caption = "Diagnosticos del " & CStr(Row.Cells("NroDocIdentidad").Value)
    End If
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub btnAceptar_Click()
   If MsgBox("¿Desea Eliminar esta Hoja?.", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
        'Elimina la HOJA completa (Cabecera - Atenciones - Diagnosticos)
        If mo_ReglasHIS.EliminarHojaHIS(oCabeceraAtencion) Then
            Call MsgBox("La Hoja se eliminó Correctamente.", vbInformation Or vbSystemModal Or vbDefaultButton1, Me.Caption)
            Unload Me
        Else
            Call MsgBox("Ocurrió un error, no se pudo eliminar.", vbCritical Or vbSystemModal Or vbDefaultButton1, Me.Caption)
        End If
   End If
End Sub

'========================================== METODOS ========================================

'CARGA DE LISTADOS DEL FORMUALRIO DE ATENCIONES
Sub CrearTablasTemp()
    'Para cargar los datos de una consulta
    With oRcs_DetalleAtencionTemp
        .Fields.Append "IdHisCabecera", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "IdHisDetalle", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "NroRegistroLote", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "NroRegistroHoja", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "IdTipoAtencion", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "DiaAtencion", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "IdHisPaciente", adVarChar, 50, adFldIsNullable + adFldUpdatable
        .Fields.Append "IdPacienteGalenHos", adVarChar, 50, adFldIsNullable + adFldUpdatable
        .Fields.Append "IdNacionalidad", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "IdTipoDocIdentidad", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "NroDocIdentidad", adVarChar, 12, adFldIsNullable + adFldUpdatable
        .Fields.Append "NroHijo", adChar, 2, adFldIsNullable + adFldUpdatable
        .Fields.Append "HC_FF_COD", adVarChar, 30, adFldIsNullable + adFldUpdatable
        .Fields.Append "IdFinanciador", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "IdDistrito", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "IdEtnia", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "TipoEdad", adVarChar, 1, adFldIsNullable + adFldUpdatable
        .Fields.Append "Edad", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "Sexo", adVarChar, 30, adFldIsNullable + adFldUpdatable
        .Fields.Append "Talla", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "Peso", adSingle, , adFldIsNullable + adFldUpdatable
        .Fields.Append "IdEstadoaEstablec", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "IdEstadoaServicio", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "IdEstado", adInteger
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
    
    'para cargar los datos de una consulta
    With oRcs_DiagnosticosTemp
        .Fields.Append "IdHisDetalleDiagnostico", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "IdHisDetalle", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "IdCIE", adInteger, 0, adFldIsNullable + adFldUpdatable
        .Fields.Append "DESCRIPCION_CIE", adVarChar, 300, adFldIsNullable + adFldUpdatable
        .Fields.Append "IdSubClasificacionDX", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "CodLAB", adVarChar, 30, adFldIsNullable + adFldUpdatable
        .Fields.Append "MSG_ALERTA", adVarChar, 60, adFldIsNullable + adFldUpdatable
        .Fields.Append "IdEstado", adInteger
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
    
    With oRcs_DetalleAtencion
        '.Fields.Append "IdHisLote", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "IdHisCabecera", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "IdHisDetalle", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "NroRegistroLote", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "NroRegistroHoja", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "IdTipoAtencion", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "DiaAtencion", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "IdHisPaciente", adVarChar, 50, adFldIsNullable + adFldUpdatable
        .Fields.Append "IdPacienteGalenHos", adVarChar, 50, adFldIsNullable + adFldUpdatable
        .Fields.Append "IdNacionalidad", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "IdTipoDocIdentidad", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "NroDocIdentidad", adVarChar, 12, adFldIsNullable + adFldUpdatable
        .Fields.Append "NroHijo", adChar, 2, adFldIsNullable + adFldUpdatable
        .Fields.Append "HC_FF_COD", adVarChar, 30, adFldIsNullable + adFldUpdatable
        .Fields.Append "IdFinanciador", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "IdDistrito", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "IdEtnia", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "TipoEdad", adVarChar, 1, adFldIsNullable + adFldUpdatable
        .Fields.Append "Edad", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "Sexo", adVarChar, 30, adFldIsNullable + adFldUpdatable
        .Fields.Append "Talla", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "Peso", adDouble, , adFldIsNullable + adFldUpdatable
        .Fields.Append "IdEstadoaEstablec", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "IdEstadoaServicio", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "IdEstado", adInteger
        
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
    Set Me.ugvResumenHIS.DataSource = oRcs_DetalleAtencion
    
    With oRcs_Diagnosticos
        .Fields.Append "IdHisDetalleDiagnostico", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "IdHisDetalle", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "IdCIE", adInteger, 0, adFldIsNullable + adFldUpdatable
        .Fields.Append "DESCRIPCION_CIE", adVarChar, 300, adFldIsNullable + adFldUpdatable
        .Fields.Append "IdSubClasificacionDX", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "CodLAB", adVarChar, 30, adFldIsNullable + adFldUpdatable
        .Fields.Append "MSG_ALERTA", adVarChar, 60, adFldIsNullable + adFldUpdatable
        .Fields.Append "IdEstado", adInteger
    
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
    Set Me.ugvDetalleDiagnosticos.DataSource = oRcs_Diagnosticos
End Sub

'CARGA DE DATOS INICIAL EN EL FORMUALRIO PRINCIPAL
Sub CargarDatosAlFormulario()
    mb_FaltaGrabarAtencion = False
    Dim oRcs_Temp As New ADODB.Recordset
    'CARGA DATOS DEL ESTABLECIMIENTO ACTUAL
    
    If mi_Opcion = sghOpciones.sghConsultar Or mi_Opcion = sghOpciones.sghModificar Or mi_Opcion = sghOpciones.sghEliminar Then
        'OBTENCION DE DATOS DE CABECERA
        Set oCabeceraAtencion = mo_ReglasHIS.ObtenerDatosCabecera(ml_IdCabeceraHIS)
        'OBTENCION DE DATOS DE MEDICO
        ml_IdMedicoResponsable = oCabeceraAtencion.IdMedico
        ml_IdEstablecimiento = oCabeceraAtencion.IdEstablecimiento
        ml_IdUsuario = oCabeceraAtencion.IdUsuario
        Dim oTablaEmpleado As New DOEmpleado
        Dim oTablaMedico As New DOMedico
        oTablaMedico.IdMedico = oCabeceraAtencion.IdMedico
        If mo_ReglasHIS.ObtenerDatosMedicoPorId(oTablaMedico, oTablaEmpleado) Then
            Me.txtResponsable.Text = oTablaEmpleado.Nombres & " " & oTablaEmpleado.ApellidoPaterno & " " & oTablaEmpleado.ApellidoMaterno
        End If
            '-------------------- INGRESO DE DATOS A FORMULARIO --------------------------
        'OBTENCION DE DATOS DEL LOTE
        mo_LoteActual.IdHisLote = oCabeceraAtencion.IdHisLote 'DEPENDE DE LA CABECERA
        mo_ReglasHIS.ObtenerDatosLotePorIdLote mo_LoteActual
        mo_cmbMes.BoundText = mo_LoteActual.Mes
        txtFechaAnio.Text = mo_LoteActual.Anio
        Me.txtPagRestante.Text = mo_LoteActual.NroHojas - mr_ReglasHIS.His_ConsultarHojasRegistradas(ml_IdEstablecimiento, mo_LoteActual.IdHisLote).RecordCount
        
        txtLote.Text = mo_LoteActual.Lote
        ml_IdLote = mo_LoteActual.IdHisLote
        txtNroPaginas.Text = mo_LoteActual.NroHojas
        Me.txtUltimaPaginaLoteActiva.Text = oCabeceraAtencion.NroHojaHis 'DEPENDE DE LA CABECERA
    End If
    
    Set oRcs_Temp = mo_ReglasHIS.HIS_DatosEstablecimientoXidEstablecimiento(ml_IdEstablecimiento)
    If oRcs_Temp.RecordCount <> 0 Then
        oRcs_Temp.MoveFirst
        Do While Not oRcs_Temp.EOF
            ml_IdDepartamentoActual = oRcs_Temp!IdDepartamento
            ms_NombreDepActual = oRcs_Temp!NombreDepartamento
            ml_IdProvinciaActual = oRcs_Temp!IdProvincia
            ms_NombreProvActual = oRcs_Temp!NombreProvincia
            ml_IdDistritoActual = oRcs_Temp!IdDistrito
            ms_NombreDistrActual = oRcs_Temp!NombreDistrito
            ml_IdEstablecimientoActual = oRcs_Temp!IdEstablecimiento
            ms_CodigoEstablecimiento = oRcs_Temp!Codigo
            ms_NombreEstablecimientoActual = oRcs_Temp!NombreEstablecimiento
            oRcs_Temp.MoveNext
        Loop
    End If
    
    'TIPO DE ACTIVIDAD
    IdTipoActividad = sghHISTipoActividad.Atencion
    
    'CARGAR DATOS DEL DIGITADOR ACTUAL
    Set oRcs_Temp = mo_ReglasHIS.ObtenerDatosDigitador(ml_IdUsuario)
    'verifica si tiene grabado el codigo de responsable de digitacion
    If IsNull(oRcs_Temp.Fields(2).Value) Then
        MsgBox "No tiene configurado el código de responsable de digitación", vbInformation, Me.Caption
        ml_CodigoResponsableDigitacion = 0
        ms_NombreRespDigitacion = oRcs_Temp.Fields(1)
    Else
        ml_CodigoResponsableDigitacion = oRcs_Temp.Fields(2)
        ms_NombreRespDigitacion = oRcs_Temp.Fields(1)
    End If
        
    'INGRESO DE VALORES CONTROLES VISUALES
    txtUbigeoDist.Text = ms_CodigoEstablecimiento
    txtUbigeoEstablecimiento.Text = ms_NombreEstablecimientoActual
    txtCodigoEstadistico.Text = ml_CodigoResponsableDigitacion & " - " & ms_NombreRespDigitacion

    If mi_Opcion = sghOpciones.sghConsultar Or mi_Opcion = sghOpciones.sghModificar Or mi_Opcion = sghOpciones.sghEliminar Then
        Set oRcs_DetalleAtencionTemp = mo_ReglasHIS.ObtenerDatosDetalleAtencion(ml_IdCabeceraHIS)
        'DETALLE ATENCION
        If oRcs_DetalleAtencionTemp.RecordCount <> 0 Then
        oRcs_DetalleAtencionTemp.MoveFirst
        Do While Not oRcs_DetalleAtencionTemp.EOF
            With oRcs_DetalleAtencion
                .AddNew
                .Fields!IdHisDetalle = oRcs_DetalleAtencionTemp!IdHisDetalle
                .Fields!IdHisCabecera = oRcs_DetalleAtencionTemp!IdHisCabecera
                .Fields!IdTipoAtencion = oRcs_DetalleAtencionTemp!IdTipoAtencion
                .Fields!NroRegistroLote = oRcs_DetalleAtencionTemp!NroRegistroLote
                .Fields!NroRegistroHoja = oRcs_DetalleAtencionTemp!NroRegistroHoja
                .Fields!DiaAtencion = oRcs_DetalleAtencionTemp!DiaAtencion
                .Fields!IdHisPaciente = oRcs_DetalleAtencionTemp!IdHisPaciente
                .Fields!IdPacienteGalenHos = oRcs_DetalleAtencionTemp!IdPacienteGalenHos
                .Fields!idnacionalidad = oRcs_DetalleAtencionTemp!idnacionalidad
                .Fields!IdTipoDocIdentidad = oRcs_DetalleAtencionTemp!IdTipoDocIdentidad
                .Fields!NroDocIdentidad = oRcs_DetalleAtencionTemp!NroDocIdentidad
                .Fields!NroHijo = oRcs_DetalleAtencionTemp!NroHijo
                .Fields!HC_FF_COD = oRcs_DetalleAtencionTemp!HC_FF_COD
                .Fields!IdFinanciador = oRcs_DetalleAtencionTemp!IdFinanciador
                .Fields!IdEtnia = oRcs_DetalleAtencionTemp!IdEtnia
                .Fields!IdDistrito = oRcs_DetalleAtencionTemp!IdDistrito
                .Fields!TipoEdad = oRcs_DetalleAtencionTemp!TipoEdad
                .Fields!Edad = oRcs_DetalleAtencionTemp!Edad
                .Fields!Sexo = oRcs_DetalleAtencionTemp!Sexo
                .Fields!Peso = oRcs_DetalleAtencionTemp!Peso
                .Fields!Talla = oRcs_DetalleAtencionTemp!Talla
                .Fields!IdEstadoaEstablec = oRcs_DetalleAtencionTemp!IdEstadoaEstablec
                .Fields!IdEstadoaServicio = oRcs_DetalleAtencionTemp!IdEstadoaServicio
                .Fields!IdEstado = 0
                .Update
            End With
            IdAtencionMax = CInt(oRcs_DetalleAtencion.Fields!IdHisDetalle) 'para obtener el ultimo id del detalle
            oRcs_DetalleAtencionTemp.MoveNext
        Loop
        End If
    End If
End Sub

Sub CargarCombosCabecera()
    mo_cmbMes.BoundColumn = "IdMes"
    mo_cmbMes.ListField = "NombreMes"
    Set mo_cmbMes.RowSource = mo_ReglasHIS.ListaMeses
    'TURNOS
    mo_cmbTurno.BoundColumn = "IdHisTurno"
    mo_cmbTurno.ListField = "Descripcion"
    Set mo_cmbTurno.RowSource = mo_ReglasHIS.ListaTurnos
    Me.cmbTurno.ListIndex = 0
    
    'SERVICIOS DE ESTABLECIMIENTO ACTUAL
    mo_cmbServicioCodigo.BoundColumn = "IdServicio"
    mo_cmbServicioCodigo.ListField = "Nombre"
    Set mo_cmbServicioCodigo.RowSource = mo_ReglasHIS.ListaServiciosPorEstablecimiento(ml_IdEstablecimiento)
    Me.cmbServicioCodigo.ListIndex = 0
End Sub

'LLENADO DE LISTADOS DE PARA EL FORMULARIO DE INGRESO DE ATENCIONES
Sub CargarCombosDetalle()
    'SI ES UN MODIFICACION O ELIMINACION, POSICIONA LOS VALORES CORRESPONDIENTES.
    If mi_Opcion <> sghAgregar Then
        mo_cmbTurno.UbicarItemDeComboBoxPorId cmbTurno, oCabeceraAtencion.IdTurno 'poner a equivalencia de su id
        mo_cmbServicioCodigo.UbicarItemDeComboBoxPorId cmbServicioCodigo, oCabeceraAtencion.IdServicio
        If oCabeceraAtencion.IdMedico <> 0 Then
            Dim oRcs_DiasProgramados As New Recordset
            Set oRcs_DiasProgramados = mo_ReglasHIS.ListarProgramacionMedicaPorMedicoYEstablecimiento(ml_IdEstablecimientoActual, oCabeceraAtencion.IdMedico, mo_cmbMes.BoundText, CInt(Me.txtFechaAnio.Text))
            If oRcs_DiasProgramados.RecordCount <> 0 Then
                oRcs_DiasProgramados.MoveFirst
                mo_Formulario.HabilitarDeshabilitar txtResponsable, False
                mo_Formulario.HabilitarDeshabilitar cmbTurno, False
                mo_Formulario.HabilitarDeshabilitar cmbServicioCodigo, False
            Else
                Call MsgBox("No tiene dias programados el Médico.", vbExclamation Or vbSystemModal, Me.Caption)
            End If
        Else
            Call MsgBox("No se encontró ID del Medico para Listar dias Programados.", vbExclamation Or vbSystemModal, Me.Caption)
        End If
    End If
    '============================= LISTADOS DE GRILLA DE ATENCIONES ===========================
    Dim oRcs_Lista As New Recordset
    
    'Tipo de Documento
    Set oRcs_Lista = mo_ReglasHIS.ListaTiposDocumento
    oRcs_Lista.MoveFirst
    mo_cmbTipoDocumento.BoundColumn = "IdDocIdentidad"
    mo_cmbTipoDocumento.ListField = "DescripcionLarga"
    Set mo_cmbTipoDocumento.RowSource = oRcs_Lista
    
    oRcs_Lista.MoveFirst
    Me.ugvResumenHIS.ValueLists.Add ("TiposDocIdentidad")
    While Not oRcs_Lista.EOF
        With Me.ugvResumenHIS.ValueLists("TiposDocIdentidad")
            .ValueListItems.Add CInt(oRcs_Lista!IdDocIdentidad), Trim(CStr(oRcs_Lista!DescripcionLarga))
            .Appearance.Font.Name = "Tahoma"
            .Appearance.Font.Size = 8
        End With
        oRcs_Lista.MoveNext
    Wend
    
    Set oRcs_Lista = Nothing
    
    'Codigo del Tipo de Financiamiento
    Set oRcs_Lista = mo_ReglasHIS.ListaFuentesFinanciamiento
    oRcs_Lista.MoveFirst
    mo_cmbFinanciador.BoundColumn = "IdCodigoFinancHis"
    mo_cmbFinanciador.ListField = "DescripcionLarga"
    Set mo_cmbFinanciador.RowSource = oRcs_Lista
    
    oRcs_Lista.MoveFirst
    Me.ugvResumenHIS.ValueLists.Add ("FuentesFinanciamiento")
    While Not oRcs_Lista.EOF
        With Me.ugvResumenHIS.ValueLists("FuentesFinanciamiento")
            .ValueListItems.Add CInt(oRcs_Lista!IdCodigoFinancHis), Trim(CStr(oRcs_Lista!DescripcionLarga))
            .Appearance.Font.Name = "Tahoma"
            .Appearance.Font.Size = 8
        End With
        oRcs_Lista.MoveNext
    Wend

    Set oRcs_Lista = Nothing
    
    'Codigo del Tipo de Etnias
    Set oRcs_Lista = mo_ReglasHIS.ListaEtnias
    oRcs_Lista.MoveFirst
    mo_cmbEtnia.BoundColumn = "codetni"
    mo_cmbEtnia.ListField = "descripcionlarga"
    Set mo_cmbEtnia.RowSource = oRcs_Lista
    oRcs_Lista.MoveFirst
    Me.ugvResumenHIS.ValueLists.Add ("Etnias")
    While Not oRcs_Lista.EOF
        With Me.ugvResumenHIS.ValueLists("Etnias")
            .ValueListItems.Add CInt(oRcs_Lista!codetni), Trim(CStr(oRcs_Lista!DescripcionLarga))
            .Appearance.Font.Name = "Tahoma"
            .Appearance.Font.Size = 8
        End With
        oRcs_Lista.MoveNext
    Wend
    
    Set oRcs_Lista = Nothing
    
    'Codigo del Tipo de Edades
    Set oRcs_Lista = mo_ReglasHIS.ListaTiposEdad
    oRcs_Lista.MoveFirst
    mo_cmbTipoEdad.BoundColumn = "IdHisTipoEdad"
    mo_cmbTipoEdad.ListField = "Descripcionlarga" '"CodigoEdad"
    Set mo_cmbTipoEdad.RowSource = oRcs_Lista
    oRcs_Lista.MoveFirst
    Me.ugvResumenHIS.ValueLists.Add ("TiposEdad")
    While Not oRcs_Lista.EOF
        With Me.ugvResumenHIS.ValueLists("TiposEdad")
            .ValueListItems.Add CInt(oRcs_Lista!IdHisTipoEdad), CStr(oRcs_Lista!CodigoEdad)
            .Appearance.Font.Name = "Tahoma"
            .Appearance.Font.Size = 8
        oRcs_Lista.MoveNext
        End With
    Wend

    Set oRcs_Lista = Nothing
    
    'Codigo del Tipo de Genero
    Set oRcs_Lista = mo_ReglasHIS.ListaTiposSexo
    oRcs_Lista.MoveFirst
    mo_cmbSexo.BoundColumn = "IdTipoSexo"
    mo_cmbSexo.ListField = "Descripcionlarga"
    Set mo_cmbSexo.RowSource = oRcs_Lista
    oRcs_Lista.MoveFirst
    Me.ugvResumenHIS.ValueLists.Add ("Genero")
    While Not oRcs_Lista.EOF
        With Me.ugvResumenHIS.ValueLists("Genero")
            .ValueListItems.Add CInt(oRcs_Lista!IdTipoSexo), CStr(oRcs_Lista!DescripcionLarga)
            .Appearance.Font.Name = "Tahoma"
            .Appearance.Font.Size = 8
        End With
        oRcs_Lista.MoveNext
    Wend

    Set oRcs_Lista = Nothing
    
    'Codigo del Tipo de Estado frente al servicio y al establecimiento
    Set oRcs_Lista = mo_ReglasHIS.ListaSituacionPaciente
    
    oRcs_Lista.MoveFirst
    mo_cmbEstadoFrenteEstablecimiento.BoundColumn = "IdTipoCondicionPaciente"
    mo_cmbEstadoFrenteEstablecimiento.ListField = "Descripcionlarga"
    Set mo_cmbEstadoFrenteEstablecimiento.RowSource = oRcs_Lista
    mo_cmbEstadoFrenteEstablecimiento.BoundText = "N"
    
    oRcs_Lista.MoveFirst
    mo_cmbEstadoFrenteServicio.BoundColumn = "IdTipoCondicionPaciente"
    mo_cmbEstadoFrenteServicio.ListField = "Descripcionlarga"
    Set mo_cmbEstadoFrenteServicio.RowSource = oRcs_Lista
    mo_cmbEstadoFrenteServicio.BoundText = "N"
    
    Me.ugvResumenHIS.ValueLists.Add ("EstadoFrenteEstablecimiento")
    Me.ugvResumenHIS.ValueLists.Add ("EstadoFrenteServicio")
    
    oRcs_Lista.MoveFirst
    While Not oRcs_Lista.EOF
        With Me.ugvResumenHIS.ValueLists("EstadoFrenteEstablecimiento")
            .ValueListItems.Add CInt(oRcs_Lista!IdTipoCondicionPaciente), Trim(CStr(oRcs_Lista!DescripcionLarga))
            .Appearance.Font.Name = "Tahoma"
            .Appearance.Font.Size = 8
        End With
        With Me.ugvResumenHIS.ValueLists("EstadoFrenteServicio")
            .ValueListItems.Add CInt(oRcs_Lista!IdTipoCondicionPaciente), Trim(CStr(oRcs_Lista!DescripcionLarga))
            .Appearance.Font.Name = "Tahoma"
            .Appearance.Font.Size = 8
        End With
        oRcs_Lista.MoveNext
    Wend

    'Listados de valores para la grilla de diagnosticos
    Set oRcs_Lista = Nothing
    
    'Codigo de Tipos de Diagnosticos
    Set oRcs_Lista = mo_ReglasHIS.ListaTiposDiagnosticos
    oRcs_Lista.MoveFirst
    Me.ugvDetalleDiagnosticos.ValueLists.Add ("ClasificacionDiagnostico")
    Me.ugvResumenDiagnosticos.ValueLists.Add ("ClasificacionDiagnostico")
    While Not oRcs_Lista.EOF
        Me.ugvDetalleDiagnosticos.ValueLists("ClasificacionDiagnostico").ValueListItems.Add CInt(oRcs_Lista!IdSubClasificacionDX), CStr(oRcs_Lista!DescripcionLarga)
        Me.ugvResumenDiagnosticos.ValueLists("ClasificacionDiagnostico").ValueListItems.Add CInt(oRcs_Lista!IdSubClasificacionDX), CStr(oRcs_Lista!DescripcionLarga)
        oRcs_Lista.MoveNext
    Wend
    oRcs_Lista.Close
    Set oRcs_Lista = Nothing
End Sub

Private Sub CargaDatosAlObjetosDeDatos()

End Sub

Private Function CargarAtencion() As Boolean
'DEPENDERA DEL TIPO DE REGISTRO - IdTipoActividad
If Not IsNull(ugvResumenHIS.ActiveRow.Cells("IdHisDetalle").Value) Then
    Select Case CInt(ugvResumenHIS.ActiveRow.Cells("IdTipoAtencion").Value)
    Case sghHISTipoActividad.Atencion
        'INGRESAR LOS DATOS A LAS CAJAS DE EDICION DEL FORMULARIO
        IdAtencion = CLng(ugvResumenHIS.ActiveRow.Cells("IdHisDetalle").Value)
        IdTipoActividad = CInt(ugvResumenHIS.ActiveRow.Cells("IdTipoAtencion").Value)
        
        ml_IdPacienteGalenHos = CLng(ugvResumenHIS.ActiveRow.Cells("IdPacienteGalenHos").Value)
        
        Me.txtNroRegistro.Text = CLng(ugvResumenHIS.ActiveRow.Cells("NroRegistroHoja").Value)
        Me.txtDia.Text = CLng(ugvResumenHIS.ActiveRow.Cells("DiaAtencion").Value)
        txtNroHC_FF_COD.Text = IIf(IsNull(ugvResumenHIS.ActiveRow.Cells("HC_FF_COD").Value), "", ugvResumenHIS.ActiveRow.Cells("HC_FF_COD").Value)
        IdCodigoNacionalidad = CStr(ugvResumenHIS.ActiveRow.Cells("IdNacionalidad").Value)
    
        Dim oRcs_Temp As New Recordset
        Set oRcs_Temp = mo_ReglasHIS.ObtenerDatosCodNacPorIdNac(IdCodigoNacionalidad)
        oRcs_Temp.MoveFirst
        txtNacionalidad.Text = CStr(oRcs_Temp!Codigo)
    
        mo_cmbTipoDocumento.UbicarItemDeComboBoxPorId cmbTipoDocumento, CStr(ugvResumenHIS.ActiveRow.Cells("IdTipoDocIdentidad").Value)
        txtNroDocumento.Text = ""
        If IsNull(ugvResumenHIS.ActiveRow.Cells("NroDocIdentidad").Value) = False Then txtNroDocumento.Text = CStr(ugvResumenHIS.ActiveRow.Cells("NroDocIdentidad").Value)
        txtOrdenFamiliar.Text = ""
        If IsNull(ugvResumenHIS.ActiveRow.Cells("NroHijo").Value) = False Then txtOrdenFamiliar.Text = ugvResumenHIS.ActiveRow.Cells("NroHijo").Value
        mo_cmbFinanciador.UbicarItemDeComboBoxPorId cmbFinanciador, CLng(ugvResumenHIS.ActiveRow.Cells("IdFinanciador").Value)
        mo_cmbEtnia.UbicarItemDeComboBoxPorId cmbEtnia, CLng(ugvResumenHIS.ActiveRow.Cells("IdEtnia").Value)
    
        ml_IdDistritoAtencion = CLng(ugvResumenHIS.ActiveRow.Cells("IdDistrito").Value)
        Dim oDistrito As New DODistrito
        oDistrito.IdDistrito = ml_IdDistritoAtencion
        mo_ReglasHIS.ConsultarDistritoPorId oDistrito
        txtDistritoProcedencia.Text = ml_IdDistritoAtencion & " - " & oDistrito.Nombre
    
        mo_cmbTipoEdad.UbicarItemDeComboBoxPorId cmbTipoEdad, CLng(ugvResumenHIS.ActiveRow.Cells("TipoEdad").Value)
        txtEdad.Text = DevuelveFormatoEdad(CStr(ugvResumenHIS.ActiveRow.Cells("Edad").Value))
        mo_cmbSexo.UbicarItemDeComboBoxPorId cmbSexo, CLng(ugvResumenHIS.ActiveRow.Cells("Sexo").Value)
        txtPeso.Text = IIf(ugvResumenHIS.ActiveRow.Cells("Peso").Value = vbNull, "", IIf(ugvResumenHIS.ActiveRow.Cells("Peso").Value = 0, "", ugvResumenHIS.ActiveRow.Cells("Peso").Value))
        txtTalla.Text = IIf(IsNull(ugvResumenHIS.ActiveRow.Cells("talla").Value), "", IIf(ugvResumenHIS.ActiveRow.Cells("talla").Value = 0, "", ugvResumenHIS.ActiveRow.Cells("talla").Value))
'        If Not ugvResumenHIS.ActiveRow.Cells("Talla").Value = vbNull Then
'           Me.txtTalla.Text = Val(ugvResumenHIS.ActiveRow.Cells("Talla").Value)
'        End If
        mo_cmbEstadoFrenteServicio.UbicarItemDeComboBoxPorId cmbEstadoFrenteServicio, CLng(ugvResumenHIS.ActiveRow.Cells("IdEstadoaServicio").Value)
        mo_cmbEstadoFrenteEstablecimiento.UbicarItemDeComboBoxPorId cmbEstadoFrenteEstablecimiento, CLng(ugvResumenHIS.ActiveRow.Cells("IdEstadoaEstablec").Value)
    
    Case sghHISTipoActividad.ActividadPreventivaPromocional, sghHISTipoActividad.ActividadConAnimales
        IdAtencion = CLng(ugvResumenHIS.ActiveRow.Cells("IdHisDetalle").Value)
        IdTipoActividad = CInt(ugvResumenHIS.ActiveRow.Cells("IdTipoAtencion").Value)
        Me.txtNroRegistro.Text = CLng(ugvResumenHIS.ActiveRow.Cells("NroRegistroHoja").Value)
        Me.txtDia.Text = CLng(ugvResumenHIS.ActiveRow.Cells("DiaAtencion").Value)
        txtNroHC_FF_COD.Text = CStr(ugvResumenHIS.ActiveRow.Cells("HC_FF_COD").Value)
        
    Case sghHISTipoActividad.ActividadMasiva
        IdAtencion = CLng(ugvResumenHIS.ActiveRow.Cells("IdHisDetalle").Value)
        IdTipoActividad = CInt(ugvResumenHIS.ActiveRow.Cells("IdTipoAtencion").Value)
        Me.txtNroRegistro.Text = CLng(ugvResumenHIS.ActiveRow.Cells("NroRegistroHoja").Value)
        Me.txtDia.Text = CLng(ugvResumenHIS.ActiveRow.Cells("DiaAtencion").Value)
        txtNroHC_FF_COD.Text = CStr(ugvResumenHIS.ActiveRow.Cells("HC_FF_COD").Value)
        txtEdad.Text = DevuelveFormatoEdad(CStr(ugvResumenHIS.ActiveRow.Cells("Edad").Value))
    End Select

    'DETALLE DIAGNOSTICOS
    If oRcs_DiagnosticosTemp.RecordCount <> 0 Then
    oRcs_DiagnosticosTemp.MoveFirst
    
    If oRcs_Diagnosticos.RecordCount <> 0 Then
        oRcs_Diagnosticos.MoveFirst
        Do While Not oRcs_Diagnosticos.EOF
            oRcs_Diagnosticos.Delete
            oRcs_Diagnosticos.MoveNext
        Loop
    End If
    
    Do While Not oRcs_DiagnosticosTemp.EOF
        With oRcs_Diagnosticos
            .AddNew
            .Fields!IdHisDetalleDiagnostico = oRcs_DiagnosticosTemp!IdHisDetalleDiagnostico
            .Fields!IdHisDetalle = oRcs_DiagnosticosTemp!IdHisDetalle
            .Fields!IdSubClasificacionDX = oRcs_DiagnosticosTemp!IdSubClasificacionDX
            .Fields!IdCIE = oRcs_DiagnosticosTemp!IdCIE
            .Fields!CodLAB = oRcs_DiagnosticosTemp!CodLAB
            .Fields!DESCRIPCION_CIE = oRcs_DiagnosticosTemp!DESCRIPCION_CIE
            .Fields!MSG_ALERTA = oRcs_DiagnosticosTemp!MSG_ALERTA
            .Fields!IdEstado = 0
            .Update
        End With
        IdDiagnosticoMax = CInt(oRcs_Diagnosticos.Fields!IdHisDetalleDiagnostico) 'para obtener el diagnostico maximo
        oRcs_DiagnosticosTemp.MoveNext
    Loop
    End If
    Set Me.ugvDetalleDiagnosticos.DataSource = oRcs_Diagnosticos
End If
End Function

Sub AdministrarKeyPreview(KeyCode As Integer)
End Sub

Sub CancelaEdicionAtencion()
    If MsgBox("Desea cancelar la edición de la atención?", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
        IniciaAtencionNueva
        IdAtencion = 0
        ml_IdPacienteGalenHos = 0
        IdTipoActividad = CInt(sghHISTipoActividad.Atencion)
        RefrescarListaAtenciones
'        Me.txtNroRegistro.SetFocus
        SeleccionaNroRegistroLibre
        Me.txtDia.SetFocus
    End If
End Sub

Sub SeleccionaNroRegistroLibre()
    Dim oRcsTemp As New ADODB.Recordset
    Dim lnIndice As Integer
    Dim RegistroLibre As Boolean
    Set oRcsTemp = mr_ReglasHIS.ObtenerDatosDetalleAtencion(IdHisCabecera)
    For lnIndice = 1 To Val(lcBuscaParametro.SeleccionaFilaParametro(272))
        RegistroLibre = True
        If oRcsTemp.RecordCount <> 0 Then
            oRcsTemp.MoveFirst
            Do While Not oRcsTemp.EOF
                If lnIndice = Int(oRcsTemp!NroRegistroHoja) Then
                    RegistroLibre = False
                End If
                oRcsTemp.MoveNext
            Loop
        End If
        If RegistroLibre = True Then
            Me.txtNroRegistro.Text = lnIndice
            Exit For
        End If
    Next
End Sub
Sub GrabaAtencionDiagnosticos()
        'Validar número de registros
        If oRcs_DetalleAtencion.RecordCount >= Val(lcBuscaParametro.SeleccionaFilaParametro(272)) Then
            MsgBox "No puede ingresar mas de " & lcBuscaParametro.SeleccionaFilaParametro(272) & " registros de atenciones", vbInformation, "HIS"
            Exit Sub
        End If
        
        'VALIDACION DE VALORES DE ATENCION
        Dim ms_Mensaje As String
        Dim ms_mensajeConsistenciaDiagnosticos As String
        ms_Mensaje = ValidarValoresAtencion
        DeshabilitaPesoTallaMayores5Anios
        
        If Len(ms_Mensaje) = 0 Then
            Me.ugvDetalleDiagnosticos.Update
            If oRcs_Diagnosticos.RecordCount <= 0 Then
                If IdTipoActividad = sghHISTipoActividad.Atencion Then
                    Call MsgBox("Ingrese los diagnósticos", vbInformation, Me.Caption)
                    Exit Sub
                End If
            Else
                oRcs_Diagnosticos.MoveFirst
                Do While Not oRcs_Diagnosticos.EOF
                    If IdTipoActividad <> sghHISTipoActividad.Atencion Then
                       If oRcs_Diagnosticos.Fields!IdSubClasificacionDX <> 102 Then
                            Call MsgBox("Cuando el codigo de activad sea un APP, AMS, AAA los diagnósticos deben ser definitivos(D)", vbInformation, Me.Caption)
                            Exit Sub
                       End If
                    End If
                    oRcs_Diagnosticos.MoveNext
                Loop
'                ms_mensajeConsistenciaDiagnosticos = mo_ReglasHIS.ValidaConsistenciaDiagnosticosHis(ml_IdLote, CInt(Me.txtDia.Text), _
'                    Val(Me.txtUltimaPaginaLoteActiva.Text), Val(Me.txtEdad.Text), Val(mo_cmbTipoEdad.BoundText), Val(mo_cmbSexo.BoundText), _
'                    Val(txtPeso.Text), Me.txtNroHC_FF_COD.Text, Val(mo_cmbEstadoFrenteEstablecimiento.BoundText), Val(mo_cmbEstadoFrenteServicio.BoundText), oRcs_Diagnosticos)
            
                ms_mensajeConsistenciaDiagnosticos = mo_ReglasHIS.ValidaConsistenciaDiagnosticosHis(ml_IdLote, CInt(Me.txtDia.Text), _
                    Val(Me.txtUltimaPaginaLoteActiva.Text), Val(Me.txtEdad.Text), ObtenerValorCombo(cmbTipoEdad.Text, "-"), ObtenerValorCombo(cmbSexo.Text, "="), _
                    Val(txtPeso.Text), Me.txtNroHC_FF_COD.Text, ObtenerValorCombo(cmbEstadoFrenteEstablecimiento.Text, "="), ObtenerValorCombo(cmbEstadoFrenteServicio.Text, "="), oRcs_Diagnosticos)
            
            End If

            If Len(ms_mensajeConsistenciaDiagnosticos) = 0 Then
'                If mb_FaltaGrabarAtencion Then
                    If MsgBox("Desea guardar la atención?", vbYesNo Or vbExclamation Or vbDefaultButton1, Me.Caption) = vbYes Then
                        If IdAtencion = 0 Then
                            AdicionAtencion
                        Else
                            ModificarAtencion
                        End If
                        IniciaAtencionNueva
                        IdAtencion = 0
                        ml_IdPacienteGalenHos = 0
                        IdTipoActividad = CInt(sghHISTipoActividad.Atencion)
                        mb_FaltaGrabarAtencion = False
                        RefrescarListaAtenciones
                        cmdIngresarDiagnosticos.Enabled = True
                        HabilitarCamposAtencionPorActividad CInt(sghHISTipoActividad.Atencion)
                        SeleccionaNroRegistroLibre
                        Me.txtDia.SetFocus
                    End If
            Else
                
                Call MsgBox(ms_mensajeConsistenciaDiagnosticos & vbCrLf & ms_Mensaje, vbInformation, "Consistencia de diagnósticos")
                'Me.ugvDetalleDiagnosticos.SetFocus
                If oRcs_Diagnosticos.RecordCount > 0 Then
                    oRcs_Diagnosticos.MoveFirst
                    ugvDetalleDiagnosticos.ActiveRow.Cells("DESCRIPCION_CIE").Selected = True
                    ugvDetalleDiagnosticos.PerformAction ssKeyActionActivateCell
                    ugvDetalleDiagnosticos.PerformAction ssKeyActionEnterEditMode
                End If
            End If
        Else
            Call MsgBox("Se encontrarón los siguientes problemas " & vbCrLf & ms_Mensaje, vbInformation, Me.Caption)
            Exit Sub
        End If
End Sub

Public Function ObtenerValorCombo(ByVal lcDescripcion As String, lcConector As String)
    Dim lnUbicacionConector As Integer
    If Trim(lcDescripcion) = "" Then
        ObtenerValorCombo = ""
    Else
        lnUbicacionConector = InStr(lcDescripcion, lcConector)
        ObtenerValorCombo = Trim(Mid(lcDescripcion, 1, lnUbicacionConector - 1))
    End If
End Function

'METODO DE ADICION DE UNA ATENCION CON SUS RESPECTIVOS DIAGNOSTICOS
Private Function AdicionAtencion() As Boolean
mb_FaltaGrabarAtencion = False
On Error GoTo AdicionAtencion_Error

If oRcs_DetalleAtencion.RecordCount > 0 Then
    oRcs_DetalleAtencion.MoveFirst
    While Not oRcs_DetalleAtencion.EOF
        oRcs_DetalleAtencion.Delete
        oRcs_DetalleAtencion.MoveNext
    Wend
End If
With oRcs_DetalleAtencion
    .AddNew
    .Fields!IdHisDetalle = IdAtencion
    .Fields!IdHisCabecera = ml_IdCabeceraHIS
    .Fields!IdTipoAtencion = IdTipoActividad
    .Fields!NroRegistroLote = 0
    .Fields!NroRegistroHoja = Val(Me.txtNroRegistro.Text)
    .Fields!DiaAtencion = Val(Me.txtDia.Text)
    .Fields!IdHisPaciente = 0
    '============================
    .Fields!IdPacienteGalenHos = ml_IdPacienteGalenHos
    '============================
    .Fields!idnacionalidad = ml_IdNacionalidadAtencion
    .Fields!IdTipoDocIdentidad = Val(mo_cmbTipoDocumento.BoundText)
    .Fields!NroDocIdentidad = Trim(Me.txtNroDocumento.Text)
    .Fields!NroHijo = Trim(Me.txtOrdenFamiliar.Text)
    .Fields!HC_FF_COD = Trim(Me.txtNroHC_FF_COD.Text)
    .Fields!IdFinanciador = Val(mo_cmbFinanciador.BoundText)
    .Fields!IdEtnia = Val(mo_cmbEtnia.BoundText)
    .Fields!IdDistrito = ml_IdDistritoAtencion
    .Fields!TipoEdad = Val(mo_cmbTipoEdad.BoundText)
    .Fields!Edad = CInt(Val(IIf(Replace(Me.txtEdad.Text, "_", "") = "", 0, Replace(Me.txtEdad.Text, "_", ""))))
    .Fields!Sexo = Val(mo_cmbSexo.BoundText)
    .Fields!Peso = Val(Me.txtPeso.Text)
    .Fields!Talla = Val(IIf(Replace(Me.txtTalla.Text, "_", "") = "", 0, Replace(Me.txtTalla.Text, "_", "")))
    .Fields!IdEstadoaEstablec = Val(mo_cmbEstadoFrenteEstablecimiento.BoundText)
    .Fields!IdEstadoaServicio = Val(mo_cmbEstadoFrenteServicio.BoundText)
    .Fields!IdEstado = 1
    .Update
End With

If mo_ReglasHIS.IngresarRegistroHIS(oCabeceraAtencion.IdHisCabecera, ml_IdUsuario, oRcs_DetalleAtencion, oRcs_Diagnosticos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, ml_IdEstablecimientoActual) Then
    AdicionAtencion = True
Else
    Call MsgBox("Ocurrió un problema con el registro de la atención.", vbCritical Or vbSystemModal, "HIS Digitación")
    AdicionAtencion = False
End If

On Error GoTo 0
Exit Function
AdicionAtencion_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure AdicionAtencion of Formulario frmFormatoHIS"
End Function

Private Sub IniciaAtencionNueva()
    On Error GoTo IniciaAtencionNueva_Error
    'LIMPIAMOS LOS CONTROLES CON SUS VALORES POR DEFECTO
    ControlesAtencionPorDefecto
    'LIMPIAMOS LA GRILLA DE DIAGNOSTICOS CON 0 FILAS
    If oRcs_DetalleAtencion.RecordCount <> 0 Then
    oRcs_DetalleAtencion.MoveFirst
    Do While Not oRcs_DetalleAtencion.EOF
        oRcs_DetalleAtencion.Delete
        oRcs_DetalleAtencion.MoveNext
    Loop
    End If
    If oRcs_Diagnosticos.RecordCount <> 0 Then
        oRcs_Diagnosticos.MoveFirst
        Do While Not oRcs_Diagnosticos.EOF
            oRcs_Diagnosticos.Delete
            oRcs_Diagnosticos.MoveNext
        Loop
    End If
    On Error GoTo 0
    Exit Sub
IniciaAtencionNueva_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IniciaAtencionNueva of Formulario frmFormatoHIS"
End Sub

'LIMPIAMOS LOS CONTROLES CON SUS VALORES POR DEFECTO
Private Sub ControlesAtencionPorDefecto()
    On Error GoTo ControlesAtencionPorDefecto_Error
    txtNroRegistro.Text = ""
    txtDia.Text = ""
    txtNroHC_FF_COD.Text = ""
    ml_IdNacionalidadAtencion = MI_IDNACIONALIDAD
    txtNacionalidad.Text = MS_NOMBRE_NAC
    mo_cmbTipoDocumento.BoundText = "1"
    txtNroDocumento.Text = ""
    txtOrdenFamiliar.Text = ""
    mo_cmbFinanciador.BoundText = "2"
    mo_cmbEtnia.BoundText = "80"
    ml_IdDistritoAtencion = ml_IdDistritoActual
    txtDistritoProcedencia.Text = ml_IdDistritoActual & " - " & ms_NombreDistrActual
    mo_cmbTipoEdad.BoundText = "1"
    txtEdad.Text = ""
    txtPeso.Text = ""
    txtTalla.Text = ""
    mo_cmbSexo.BoundText = "1"
    mo_cmbEstadoFrenteEstablecimiento.BoundText = "1"
    mo_cmbEstadoFrenteServicio.BoundText = "1"
    On Error GoTo 0
    Exit Sub
ControlesAtencionPorDefecto_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ControlesAtencionPorDefecto of Formulario frmMantenimientoHIS"
End Sub

Private Function ModificarAtencion() As Boolean
mb_FaltaGrabarAtencion = False
On Error GoTo ModificarAtencion_Error

If oRcs_DetalleAtencion.RecordCount > 0 Then
    oRcs_DetalleAtencion.MoveFirst
    While Not oRcs_DetalleAtencion.EOF
        oRcs_DetalleAtencion.Delete
        oRcs_DetalleAtencion.MoveNext
    Wend
End If
With oRcs_DetalleAtencion
    .AddNew
    .Fields!IdHisDetalle = IdAtencion
    .Fields!IdHisCabecera = 0
    .Fields!IdTipoAtencion = IdTipoActividad
    .Fields!DiaAtencion = Trim(Me.txtDia.Text)
    .Fields!IdHisPaciente = 0
    .Fields!IdPacienteGalenHos = ml_IdPacienteGalenHos
    .Fields!idnacionalidad = ml_IdNacionalidadAtencion
    .Fields!IdTipoDocIdentidad = Val(mo_cmbTipoDocumento.BoundText)
    .Fields!NroDocIdentidad = Trim(Me.txtNroDocumento.Text)
    .Fields!NroHijo = Trim(Me.txtOrdenFamiliar.Text)
    .Fields!HC_FF_COD = Trim(Me.txtNroHC_FF_COD.Text)
    .Fields!IdFinanciador = Val(mo_cmbFinanciador.BoundText)
    .Fields!IdEtnia = Val(mo_cmbEtnia.BoundText)
    .Fields!IdDistrito = ml_IdDistritoAtencion
    .Fields!TipoEdad = Val(mo_cmbTipoEdad.BoundText)
    .Fields!Edad = CInt(Val(IIf(Replace(Me.txtEdad.Text, "_", "") = "", 0, Replace(Me.txtEdad.Text, "_", ""))))
    .Fields!Sexo = Val(mo_cmbSexo.BoundText)
    .Fields!Peso = Val(Me.txtPeso.Text)
    .Fields!Talla = Val(IIf(Replace(Me.txtTalla.Text, "_", "") = "", 0, Replace(Me.txtTalla.Text, "_", "")))
    .Fields!IdEstadoaEstablec = Val(mo_cmbEstadoFrenteEstablecimiento.BoundText)
    .Fields!IdEstadoaServicio = Val(mo_cmbEstadoFrenteServicio.BoundText)
    .Fields!NroRegistroLote = 0
    .Fields!NroRegistroHoja = Val(txtNroRegistro.Text)
    .Fields!IdEstado = 2
    .Update
End With

If mo_ReglasHIS.ActualizarRegistroHIS(oCabeceraAtencion.IdHisCabecera, ml_IdUsuario, oRcs_DetalleAtencion, oRcs_Diagnosticos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, ml_IdEstablecimientoActual) Then
    ModificarAtencion = True
Else
    Call MsgBox("Ocurrió un problema con la actualización de la atención.", vbCritical, "HIS Digitación")
    ModificarAtencion = False
End If

On Error GoTo 0
Exit Function
ModificarAtencion_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ModificarAtencion of Formulario frmFormatoHIS"
End Function

Private Function EliminarAtencion() As Boolean
    mb_FaltaGrabarAtencion = False
    On Error GoTo ModificarAtencion_Error
    
    'Ingresamos el valor del ID en la Grilla
    If Not IsNull(Me.ugvResumenHIS.ActiveRow.Cells("IdHisDetalle").Value) Then
        IdAtencion = CLng(Me.ugvResumenHIS.ActiveRow.Cells("IdHisDetalle").Value)
        If mo_ReglasHIS.EliminarRegistroHIS(IdAtencion, ml_IdUsuario, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc) Then
            EliminarAtencion = True
            IdAtencion = 0
        Else
            Call MsgBox("Ocurrió un problema con la eliminación de la atención.", vbCritical, Me.Caption)
            EliminarAtencion = False
        End If
    End If
    On Error GoTo 0
    Exit Function
ModificarAtencion_Error:
    IdAtencion = 0
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ModificarAtencion of Formulario frmFormatoHIS"
End Function

Private Sub AdicionDiagnostico()
    On Error GoTo AdicionDiagnostico_Error
    If mi_Opcion = sghAgregar Then
        If mb_SeleccionoLote = False Then
            Call MsgBox("Debe ingresar previamente el lote", vbInformation, Me.Caption)
            Exit Sub
        End If
        If mb_SeleccionoHoja = False Then
            Call MsgBox("Debe ingresar previamente la hoja", vbInformation, Me.Caption)
            Exit Sub
        End If
        If mb_SeleccionoMedico = False Then
            Call MsgBox("Debe ingresar previamente el responsable de atención", vbInformation, Me.Caption)
            Exit Sub
        End If
    End If
    
    If oRcs_Diagnosticos.RecordCount < 6 Then
        Dim lbIngresarDx As Boolean
        Dim lnIdCIE As Long
        Dim lcNombreDx As String
        Dim lnMasDeUnDx As Integer
        Dim oBusqueda As New SIGHhisDigitacion.BusquedaProductosHis
        Dim oDoFACTCATALOGOSERVICIOS As New sighcomun.DOHis_FactCatalogoServicios
        oBusqueda.CodigoDx = ""
        lnIdCIE = 0
        lcNombreDx = ""
        lnMasDeUnDx = 0
        lbIngresarDx = False
        oBusqueda.MostrarFormulario
        If oBusqueda.BotonPresionado = sghAceptar Then
            lbIngresarDx = True
            lnIdCIE = oBusqueda.IdRegistroSeleccionado
            lcNombreDx = oBusqueda.descripciondiagcpt
            lnMasDeUnDx = oBusqueda.MasDeUnDiagnosticos
            'VALIDACION DE DIAGNOSTICOS REPETIDOS
            If oRcs_Diagnosticos.RecordCount <> 0 Then
            'SE CLONARA PARA BUSCAR EL DATO
                oRcs_Diagnosticos.MoveFirst
                Do While Not oRcs_Diagnosticos.EOF
                    If Not IsNull(oRcs_Diagnosticos!IdCIE) Then 'indica si es el primero ingresado
                        If lnIdCIE = CLng(oRcs_Diagnosticos!IdCIE) Then
                            If lnMasDeUnDx = 0 Then
                                lbIngresarDx = False
                                Call MsgBox("El Producto His ya fue ingresado.", vbExclamation, Me.Caption)
                            End If
                        End If
                    End If
                    oRcs_Diagnosticos.MoveNext
                Loop
                oRcs_Diagnosticos.MoveFirst
            End If
        ElseIf oBusqueda.BotonPresionado = sghCancelar Then
            Exit Sub
        End If
        Set oBusqueda = Nothing
        If lbIngresarDx = True Then
            With oRcs_Diagnosticos
                .AddNew
                .Fields!IdCIE = lnIdCIE
                .Fields!IdHisDetalleDiagnostico = IdDetalleDiagnostico
                .Fields!IdHisDetalle = IdAtencion
                If IdTipoActividad = sghHISTipoActividad.Atencion Then
                    'En caso que sea una atencion
                    .Fields!IdSubClasificacionDX = 101
                Else
                    'En caso que sea una atencion
                    .Fields!IdSubClasificacionDX = 102
                End If
                .Fields!DESCRIPCION_CIE = lcNombreDx
                .Fields!CodLAB = ""
                .Fields!IdEstado = 1
                .Update
            End With
        End If
    End If
    Set ugvDetalleDiagnosticos.DataSource = oRcs_Diagnosticos
    Me.ugvDetalleDiagnosticos.ActiveRow.Activation = ssActivationAllowEdit
    ugvDetalleDiagnosticos.SetFocus
    ugvDetalleDiagnosticos.ActiveRow.Cells("DESCRIPCION_CIE").Selected = True
    ugvDetalleDiagnosticos.PerformAction ssKeyActionActivateCell
    ugvDetalleDiagnosticos.PerformAction ssKeyActionEnterEditMode
    On Error GoTo 0
    Exit Sub
AdicionDiagnostico_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure AdicionDiagnostico of Formulario frmFormatoHIS"
End Sub

Private Sub EliminarDiagnostico()
On Error GoTo EliminarDiagnostico_Error

If Me.ugvDetalleDiagnosticos.ActiveRow Is Nothing Then
   Call MsgBox("Seleccione el diagnóstico a eliminar", vbExclamation Or vbSystemModal, "HIS Digitación")
    Exit Sub
End If
    
        With oRcs_Diagnosticos
            If Not .EOF And Not .BOF Then
                .Delete
                .Update
            End If
        End With

On Error GoTo 0
Exit Sub
EliminarDiagnostico_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure EliminarDetalleDiagnostico of Formulario frmFormatoHIS"
End Sub

'=================================== METODOS AUXILIARES ========================================

'ANALIZA EL TIPO DE ACTIVIDAD QUE CORRESPONDE AL CONTENIDO EN EL CAMPO DE HC_FF_COD
Function ObtenerTipoActividad() As Integer
Dim mi_existe As Integer

mi_existe = InStr(UCase(Trim(Me.txtNroHC_FF_COD.Text)), "APP")
If mi_existe > 0 Then
    IdTipoActividad = sghHISTipoActividad.ActividadPreventivaPromocional
    ObtenerTipoActividad = IdTipoActividad
    Exit Function
End If

mi_existe = InStr(UCase(Trim(Me.txtNroHC_FF_COD.Text)), "AMS")
If mi_existe > 0 Then
    IdTipoActividad = sghHISTipoActividad.ActividadMasiva
    ObtenerTipoActividad = IdTipoActividad
    Exit Function
End If

mi_existe = InStr(UCase(Trim(Me.txtNroHC_FF_COD.Text)), "AAA")
If mi_existe > 0 Then
    IdTipoActividad = sghHISTipoActividad.ActividadConAnimales
    ObtenerTipoActividad = IdTipoActividad
    Exit Function
End If
End Function

'HABILITA LOS CAMPOS DE LA ACTIVIDADES DEPENDIENDO DE SU TIPO
Private Sub HabilitarCamposAtencionPorActividad(IdCodAct As Integer)
Dim Habilitado As Boolean
Select Case IdCodAct
    Case sghHISTipoActividad.Atencion
        Habilitado = True
'        mo_Formulario.HabilitarDeshabilitar txtNroRegistro, Habilitado
'        mo_Formulario.HabilitarDeshabilitar txtDia, Habilitado
        mo_Formulario.HabilitarDeshabilitar txtNacionalidad, Habilitado
        mo_Formulario.HabilitarDeshabilitar cmbTipoDocumento, Habilitado
        mo_Formulario.HabilitarDeshabilitar txtNroDocumento, Habilitado
        If mo_cmbTipoDocumento.BoundText = 8 Then
            mo_Formulario.HabilitarDeshabilitar txtOrdenFamiliar, Habilitado
        Else
            mo_Formulario.HabilitarDeshabilitar txtOrdenFamiliar, False
        End If
        mo_Formulario.HabilitarDeshabilitar cmbFinanciador, Habilitado
        mo_Formulario.HabilitarDeshabilitar cmbEtnia, Habilitado
        mo_Formulario.HabilitarDeshabilitar txtDistritoProcedencia, Habilitado
        mo_Formulario.HabilitarDeshabilitar cmbTipoEdad, Habilitado
        mo_Formulario.HabilitarDeshabilitar txtEdad, Habilitado
        mo_Formulario.HabilitarDeshabilitar cmbSexo, Habilitado
        If lcBuscaParametro.SeleccionaFilaParametro(329) = "S" Then
            mo_Formulario.HabilitarDeshabilitar Me.txtPeso, Habilitado
            mo_Formulario.HabilitarDeshabilitar Me.txtTalla, Habilitado
        Else
            mo_Formulario.HabilitarDeshabilitar Me.txtPeso, False
            mo_Formulario.HabilitarDeshabilitar Me.txtTalla, False
        End If
        mo_Formulario.HabilitarDeshabilitar cmbEstadoFrenteEstablecimiento, Habilitado
        mo_Formulario.HabilitarDeshabilitar cmbEstadoFrenteServicio, Habilitado
'        Me.txtNacionalidad.SetFocus
        
    Case sghHISTipoActividad.ActividadPreventivaPromocional, sghHISTipoActividad.ActividadConAnimales
        Habilitado = False
'        mo_Formulario.HabilitarDeshabilitar txtNroRegistro, Habilitado
'        mo_Formulario.HabilitarDeshabilitar txtDia, Habilitado
        mo_Formulario.HabilitarDeshabilitar txtNacionalidad, Habilitado
        mo_Formulario.HabilitarDeshabilitar cmbTipoDocumento, Habilitado
        txtNroDocumento.Text = ""
        mo_Formulario.HabilitarDeshabilitar txtNroDocumento, Habilitado
        txtOrdenFamiliar.Text = ""
        mo_Formulario.HabilitarDeshabilitar txtOrdenFamiliar, Habilitado
        mo_Formulario.HabilitarDeshabilitar cmbFinanciador, Habilitado
        mo_Formulario.HabilitarDeshabilitar cmbEtnia, Habilitado
        mo_Formulario.HabilitarDeshabilitar txtDistritoProcedencia, Habilitado
        mo_Formulario.HabilitarDeshabilitar cmbTipoEdad, Habilitado
        txtEdad.Text = ""
        mo_Formulario.HabilitarDeshabilitar txtEdad, Habilitado
        mo_Formulario.HabilitarDeshabilitar cmbSexo, Habilitado
        txtPeso.Text = ""
        mo_Formulario.HabilitarDeshabilitar txtPeso, Habilitado
        txtTalla.Text = ""
        mo_Formulario.HabilitarDeshabilitar txtTalla, Habilitado
        mo_Formulario.HabilitarDeshabilitar cmbEstadoFrenteEstablecimiento, Habilitado
        mo_Formulario.HabilitarDeshabilitar cmbEstadoFrenteServicio, Habilitado
        Me.cmdIngresarDiagnosticos.Enabled = True
        
    Case sghHISTipoActividad.ActividadMasiva
        Habilitado = False
'        mo_Formulario.HabilitarDeshabilitar txtNroRegistro, Habilitado
'        mo_Formulario.HabilitarDeshabilitar txtDia, Habilitado
        mo_Formulario.HabilitarDeshabilitar txtNacionalidad, Habilitado
        mo_Formulario.HabilitarDeshabilitar cmbTipoDocumento, Habilitado
        txtNroDocumento.Text = ""
        mo_Formulario.HabilitarDeshabilitar txtNroDocumento, Habilitado
        txtOrdenFamiliar.Text = ""
        mo_Formulario.HabilitarDeshabilitar txtOrdenFamiliar, Habilitado
        mo_Formulario.HabilitarDeshabilitar cmbFinanciador, Habilitado
        mo_Formulario.HabilitarDeshabilitar cmbEtnia, Habilitado
        mo_Formulario.HabilitarDeshabilitar txtDistritoProcedencia, Habilitado
        mo_Formulario.HabilitarDeshabilitar cmbTipoEdad, Not (Habilitado)
        mo_Formulario.HabilitarDeshabilitar txtEdad, Not (Habilitado)
        mo_Formulario.HabilitarDeshabilitar cmbSexo, Habilitado
        txtPeso.Text = ""
        mo_Formulario.HabilitarDeshabilitar txtPeso, Habilitado
        txtTalla.Text = ""
        mo_Formulario.HabilitarDeshabilitar txtTalla, Habilitado
        mo_Formulario.HabilitarDeshabilitar cmbEstadoFrenteEstablecimiento, Habilitado
        mo_Formulario.HabilitarDeshabilitar cmbEstadoFrenteServicio, Habilitado
'        Me.txtEdad.SetFocus
    End Select
End Sub

Private Sub BloquearControlesAtencion()
'deshabilitar controles de atencion para que el usuario confirme su actualizacion
    mo_Formulario.HabilitarDeshabilitar Me.txtNroRegistro, False
    mo_Formulario.HabilitarDeshabilitar Me.txtDia, False
    mo_Formulario.HabilitarDeshabilitar Me.txtNroHC_FF_COD, False
    mo_Formulario.HabilitarDeshabilitar txtNacionalidad, False
    mo_Formulario.HabilitarDeshabilitar cmbTipoDocumento, False
    mo_Formulario.HabilitarDeshabilitar txtNroDocumento, False
    mo_Formulario.HabilitarDeshabilitar txtOrdenFamiliar, False
    mo_Formulario.HabilitarDeshabilitar cmbFinanciador, False
    mo_Formulario.HabilitarDeshabilitar cmbEtnia, False
    mo_Formulario.HabilitarDeshabilitar txtDistritoProcedencia, False
    mo_Formulario.HabilitarDeshabilitar cmbTipoEdad, False
    mo_Formulario.HabilitarDeshabilitar txtEdad, False
    mo_Formulario.HabilitarDeshabilitar cmbSexo, False
    mo_Formulario.HabilitarDeshabilitar txtPeso, False
    mo_Formulario.HabilitarDeshabilitar txtTalla, False
    mo_Formulario.HabilitarDeshabilitar cmbEstadoFrenteEstablecimiento, False
    mo_Formulario.HabilitarDeshabilitar cmbEstadoFrenteServicio, False
End Sub

Private Function ValidarValoresAtencion() As String
Dim ms_ValidacionAtencion As String
Dim mb_FaltaDato As Boolean
mb_FaltaDato = False
'LIMPIAMOS LOS MARCADORES DE ERRORES
If IdTipoActividad = sghHISTipoActividad.Atencion Then
    With mo_Formulario
        .HabilitarDeshabilitar txtNroRegistro, True
        .HabilitarDeshabilitar txtDia, True
        .HabilitarDeshabilitar txtNroHC_FF_COD, True
        .HabilitarDeshabilitar txtNacionalidad, True
        .HabilitarDeshabilitar cmbTipoDocumento, True
        .HabilitarDeshabilitar txtNroDocumento, True
        .HabilitarDeshabilitar txtOrdenFamiliar, True
        .HabilitarDeshabilitar cmbFinanciador, True
        .HabilitarDeshabilitar cmbEtnia, True
        .HabilitarDeshabilitar txtDistritoProcedencia, True
        .HabilitarDeshabilitar cmbTipoEdad, True
        .HabilitarDeshabilitar txtEdad, True
        .HabilitarDeshabilitar cmbSexo, True
        If lcBuscaParametro.SeleccionaFilaParametro(329) = "S" Then
            .HabilitarDeshabilitar txtPeso, True
            .HabilitarDeshabilitar txtTalla, True
        Else
            .HabilitarDeshabilitar txtPeso, False
            .HabilitarDeshabilitar txtTalla, False
        End If
        .HabilitarDeshabilitar cmbEstadoFrenteServicio, True
        .HabilitarDeshabilitar cmbEstadoFrenteEstablecimiento, True
    End With
Else
    With mo_Formulario
        .HabilitarDeshabilitar txtNroRegistro, True
        .HabilitarDeshabilitar txtDia, True
        .HabilitarDeshabilitar txtNroHC_FF_COD, True
    End With
End If

'Validar cabecera
If Trim(Me.txtLote.Text) = "" Then
    ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- No se dígito el lote"
    If mb_FaltaDato = False Then
        Me.txtLote.SetFocus
        mb_FaltaDato = True
    End If
    Me.txtLote.BackColor = ml_ColorError
End If
If Trim(Me.txtResponsable.Text) = "" Then
    ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- No se dígito el responsable"
    If mb_FaltaDato = False Then
        Me.txtResponsable.SetFocus
        mb_FaltaDato = True
    End If
    Me.txtResponsable.BackColor = ml_ColorError
End If
 
'VALIDACION DE LOS CAMPOS INGRESADOS ANTES DE EDITAR LOS DIAGNOSTICOS

If Trim(Me.txtNroRegistro.Text) = "" Then
    ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- No se dígito el número de registro"
    If mb_FaltaDato = False Then
        Me.txtNroRegistro.SetFocus
        mb_FaltaDato = True
    End If
    Me.txtNroRegistro.BackColor = ml_ColorError
Else
    If CInt(Me.txtNroRegistro.Text) > Val(lcBuscaParametro.SeleccionaFilaParametro(272)) Then
        ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- El nro de registro no puede ser mayor a " & lcBuscaParametro.SeleccionaFilaParametro(272)
        If mb_FaltaDato = False Then
            Me.txtNroRegistro.SetFocus
            mb_FaltaDato = True
        End If
        Me.txtNroRegistro.BackColor = ml_ColorError
    Else
        If IdAtencion = 0 Then
        Dim oRcsTemp As New ADODB.Recordset
        Dim RegistroUsado As Boolean
        Set oRcsTemp = mr_ReglasHIS.ObtenerDatosDetalleAtencion(ml_IdCabeceraHIS)
        RegistroUsado = False
        If oRcsTemp.RecordCount <> 0 Then
            oRcsTemp.MoveFirst
            Do While Not oRcsTemp.EOF
                If CInt(Me.txtNroRegistro.Text) = Int(oRcsTemp!NroRegistroHoja) Then
                    RegistroUsado = True
                End If
                oRcsTemp.MoveNext
            Loop
        End If
        If RegistroUsado = True Then
                ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- El nro de registro de atención ya fue digitado"
                If mb_FaltaDato = False Then
                    Me.txtNroRegistro.SetFocus
                    mb_FaltaDato = True
                End If
                Me.txtNroRegistro.BackColor = ml_ColorError
        End If
        End If
    End If
End If

If Trim(Me.txtDia.Text) = "" Then
    ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- No se dígito el día de atención"
    If mb_FaltaDato = False Then
        Me.txtDia.SetFocus
        mb_FaltaDato = True
    End If
    Me.txtDia.BackColor = ml_ColorError
Else
    Dim oRcs_DetalleProgramacionTemp As New ADODB.Recordset
    Dim DiaValido As Boolean
    DiaValido = False
    Set oRcs_DetalleProgramacionTemp = mr_ReglasHIS.ObtenerDatosProgramacionMedica(ml_IdEstablecimientoActual, mo_cmbServicioCodigo.BoundText, ml_IdMedicoResponsable, CLng(txtFechaAnio.Text), CLng(mo_cmbMes.BoundText), mo_cmbTurno.BoundText)
    
    If oRcs_DetalleProgramacionTemp.RecordCount <> 0 Then
        oRcs_DetalleProgramacionTemp.MoveFirst
        Do While Not oRcs_DetalleProgramacionTemp.EOF
            If Day(CDate(oRcs_DetalleProgramacionTemp!FechaProgramada)) = CInt(Me.txtDia.Text) Then
                DiaValido = True
            End If
            oRcs_DetalleProgramacionTemp.MoveNext
        Loop
    End If
    If DiaValido = False Then
            ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- El día ingresado no es válido"
            If mb_FaltaDato = False Then
                Me.txtDia.SetFocus
                mb_FaltaDato = True
            End If
            Me.txtDia.BackColor = ml_ColorError
    End If
End If

If Trim(Me.txtNroHC_FF_COD.Text) = "" Then
    ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- No se dígito el código HC/FF/Cod. atención"
    If mb_FaltaDato = False Then
        Me.txtNroHC_FF_COD.SetFocus
        mb_FaltaDato = True
    End If
    Me.txtNroHC_FF_COD.BackColor = ml_ColorError
End If

If IdTipoActividad = sghHISTipoActividad.Atencion Then
    'NACIONALIDAD
    If Trim(txtNacionalidad.Text) = "" Then
        ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- Ingrese el código de la nacionalidad"
        If mb_FaltaDato = False Then
            Me.txtNacionalidad.SetFocus
            mb_FaltaDato = True
        End If
        Me.txtNacionalidad.BackColor = ml_ColorError
    Else
        Dim oRcs_Temp As New Recordset
        Set oRcs_Temp = mo_ReglasHIS.ObtenerDatosCodNacPorCodigo(txtNacionalidad.Text)
        If oRcs_Temp.RecordCount > 0 Then
            oRcs_Temp.MoveFirst
            IdCodigoNacionalidad = Val(oRcs_Temp!IdPais)
        Else
            ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- El código de la nacionalidad es incorrecto, oprima F11 para verificar"
            If mb_FaltaDato = False Then
                Me.txtNacionalidad.SetFocus
                mb_FaltaDato = True
            End If
            Me.txtNacionalidad.BackColor = ml_ColorError
        End If
    End If
    
    If Trim(Me.cmbTipoDocumento.Text) = "" Then
        ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- Seleccione el tipo de documento"
        If mb_FaltaDato = False Then
            Me.cmbTipoDocumento.SetFocus
            mb_FaltaDato = True
        End If
        Me.cmbTipoDocumento.BackColor = ml_ColorError
    Else
        If IdCodigoNacionalidad = 166 Then
            If mo_cmbTipoDocumento.BoundText <> 1 And mo_cmbTipoDocumento.BoundText <> 8 Then
                ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- El tipo de documento de identidad para Perú solo puede ser 1 ó 8"
                If mb_FaltaDato = False Then
                    Me.cmbTipoDocumento.SetFocus
                    mb_FaltaDato = True
                End If
                Me.cmbTipoDocumento.BackColor = ml_ColorError
            End If
        End If
    End If
    
    If lcBuscaParametro.SeleccionaFilaParametro(340) = "S" Then
        If Trim(Me.txtNroDocumento.Text) = "" Then
            ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- No se dígito el número de documento"
            If mb_FaltaDato = False Then
                Me.txtNroDocumento.SetFocus
                mb_FaltaDato = True
            End If
            Me.txtNroDocumento.BackColor = ml_ColorError
        Else
            If mo_cmbTipoDocumento.BoundText = 1 Or mo_cmbTipoDocumento.BoundText = 8 Then
                If Len(Trim(Me.txtNroDocumento.Text)) <> 8 Then
                    ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- El DNI debe tener 8 dígitos"
                    If mb_FaltaDato = False Then
                        Me.txtNroDocumento.SetFocus
                        mb_FaltaDato = True
                    End If
                    Me.txtNroDocumento.BackColor = ml_ColorError
                End If
            End If
        End If
    End If
    
    If mo_cmbTipoDocumento.BoundText = 8 Then
        If Trim(Me.txtOrdenFamiliar.Text) = "" Then
            ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- No se dígito el número de hijo"
            If mb_FaltaDato = False Then
                Me.txtOrdenFamiliar.SetFocus
                mb_FaltaDato = True
            End If
            Me.txtOrdenFamiliar.BackColor = ml_ColorError
        End If
    Else
        mo_Formulario.HabilitarDeshabilitar txtOrdenFamiliar, False
    End If
    
    If Trim(Me.cmbFinanciador.Text) = "" Then
        ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- No se dígito el financiador de la atención"
        If mb_FaltaDato = False Then
            Me.cmbFinanciador.SetFocus
            mb_FaltaDato = True
        End If
        Me.cmbFinanciador.BackColor = ml_ColorError
    End If
    
    If Trim(Me.cmbEtnia.Text) = "" Then
        ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- No se dígito la etnia"
        If mb_FaltaDato = False Then
            Me.cmbEtnia.SetFocus
            mb_FaltaDato = True
        End If
        Me.cmbEtnia.BackColor = ml_ColorError
    End If
    
    'ALTERO EL CONTENIDO DE DISTRITO DE LA PROCEDENCIA
    If ml_IdDistritoAtencion = 0 Then
        ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- No se busco o se alteró el distrito origen"
        If mb_FaltaDato = False Then
            Me.txtDistritoProcedencia.SetFocus
            mb_FaltaDato = True
        End If
        Me.txtDistritoProcedencia.BackColor = ml_ColorError
    End If
    
    'EDAD Y TIPO DE EDAD
    If CInt(IIf(Replace(Me.txtEdad.Text, "_", "") = "", 0, Replace(Me.txtEdad.Text, "_", ""))) = 0 Then
        ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- No se dígito la edad"
        If mb_FaltaDato = False Then
            Me.txtEdad.SetFocus
            mb_FaltaDato = True
        End If
        Me.txtEdad.BackColor = ml_ColorError
    Else
        If CInt(Val(mo_cmbTipoEdad.BoundText)) = sghHISTipoEdades.Dias And CInt(IIf(Replace(Me.txtEdad.Text, "_", "") = "", 0, Replace(Me.txtEdad.Text, "_", ""))) > 30 Then
            ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- La edad maxima en días es de 30"
            If mb_FaltaDato = False Then
                Me.txtEdad.SetFocus
                mb_FaltaDato = True
            End If
            Me.txtEdad.BackColor = ml_ColorError
            Me.cmbTipoEdad.BackColor = ml_ColorError
        ElseIf CInt(Val(mo_cmbTipoEdad.BoundText)) = sghHISTipoEdades.Meses And CInt(IIf(Replace(Me.txtEdad.Text, "_", "") = "", 0, Replace(Me.txtEdad.Text, "_", ""))) > 11 Then
           
            ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- La edad maxima en meses es de 11"
            If mb_FaltaDato = False Then
                Me.txtEdad.SetFocus
                mb_FaltaDato = True
            End If
            Me.txtEdad.BackColor = ml_ColorError
            Me.cmbTipoEdad.BackColor = ml_ColorError
            
        ElseIf CInt(Val(mo_cmbTipoEdad.BoundText)) = sghHISTipoEdades.Años And CInt(IIf(Replace(Me.txtEdad.Text, "_", "") = "", 0, Replace(Me.txtEdad.Text, "_", ""))) > 99 Then
            ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- La edad maxima en años es de 99"
            If mb_FaltaDato = False Then
                Me.txtEdad.SetFocus
                mb_FaltaDato = True
            End If
            Me.txtEdad.BackColor = ml_ColorError
            Me.cmbTipoEdad.BackColor = ml_ColorError
        End If
    End If

    
    If mb_PesoTallaHabilitados = True Then
        'PESO
        If lcBuscaParametro.SeleccionaFilaParametro(329) = "S" Then
           If IsNumeric(Me.txtPeso.Text) Then
               If CSng(Me.txtPeso.Text) > 300 Or CSng(Me.txtPeso.Text) = 0 Then
                   ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- El peso no debe ser mayor a 300 Kg."
                    If mb_FaltaDato = False Then
                        Me.txtPeso.SetFocus
                        mb_FaltaDato = True
                    End If
                   txtPeso.BackColor = ml_ColorError
               End If
           Else
               ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- El peso no tiene formato número"
                If mb_FaltaDato = False Then
                    Me.txtPeso.SetFocus
                    mb_FaltaDato = True
                End If
               txtPeso.BackColor = ml_ColorError
           End If
           'TALLA
           If CInt(IIf(Replace(Me.txtTalla.Text, "_", "") = "", 0, Replace(Me.txtTalla.Text, "_", ""))) > 210 Or CInt(IIf(Replace(Me.txtTalla.Text, "_", "") = "", 0, Replace(Me.txtTalla.Text, "_", ""))) = 0 Then
               ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- La talla menor a 2.10 mts."
                If mb_FaltaDato = False Then
                    Me.txtTalla.SetFocus
                    mb_FaltaDato = True
                End If
               txtTalla.BackColor = ml_ColorError
           End If
        End If
    End If
    
    If IdAtencion = 0 Then
        If mo_cmbTipoDocumento.BoundText = 8 Then
             If Trim(Me.txtNroDocumento.Text) <> "" And Trim(Me.txtOrdenFamiliar.Text) <> "" Then
                If oRcs_DetalleAtencion.RecordCount <> 0 Then
                    oRcs_DetalleAtencion.MoveFirst
                    Do While Not oRcs_DetalleAtencion.EOF
                        If mo_cmbTipoDocumento.BoundText = oRcs_DetalleAtencion!IdTipoDocIdentidad And _
                            Trim(Me.txtNroDocumento.Text) = Trim(oRcs_DetalleAtencion!NroDocIdentidad) And _
                            Trim(Me.txtOrdenFamiliar.Text) = Trim(oRcs_DetalleAtencion!NroHijo) Then
                                ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- El paciente ya fue digitado en esta hoja"
                                If mb_FaltaDato = False Then
                                    Me.txtNroDocumento.SetFocus
                                    mb_FaltaDato = True
                                End If
                                Exit Do
                        End If
                        oRcs_DetalleAtencion.MoveNext
                    Loop
                    oRcs_DetalleAtencion.MoveFirst
                End If
             End If
        Else
        'COMENTADO POR FRANKLIN - MANUEL
'             If Trim(Me.txtNroDocumento.Text) <> "" Then
'                If oRcs_DetalleAtencion.RecordCount <> 0 Then
'                    oRcs_DetalleAtencion.MoveFirst
'                    Do While Not oRcs_DetalleAtencion.EOF
'                        If mo_cmbTipoDocumento.BoundText = oRcs_DetalleAtencion!IdTipoDocIdentidad And _
'                            Trim(Me.txtNroDocumento.Text) = Trim(oRcs_DetalleAtencion!NroDocIdentidad) Then
'                                ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- El paciente ya fue digitado en esta hoja"
'                                Exit Do
'                        End If
'                        oRcs_DetalleAtencion.MoveNext
'                    Loop
'                    oRcs_DetalleAtencion.MoveFirst
'                End If
'            End If
        End If
    End If
    
    'ESTADO DEL SERVICIO Y EL ESTABLECIMIENTO
    Dim mi_EstadoServicio As Integer
    Dim mi_EstadoEstablecimiento As Integer
    mi_EstadoServicio = CInt(Val(mo_cmbEstadoFrenteServicio.BoundText))
    mi_EstadoEstablecimiento = CInt(Val(mo_cmbEstadoFrenteEstablecimiento.BoundText))
    
    'NUEVOS
    If mi_EstadoEstablecimiento = sghHISEstados.Nuevo Then
        If mi_EstadoServicio <> sghHISEstados.Nuevo Then
            ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- Si es nuevo en el establecimiento, debe ser nuevo a los servicios"
            If mb_FaltaDato = False Then
                Me.cmbEstadoFrenteServicio.SetFocus
                mb_FaltaDato = True
            End If
            Me.cmbEstadoFrenteServicio.BackColor = ml_ColorError
            Me.cmbEstadoFrenteEstablecimiento.BackColor = ml_ColorError
        End If
        
    'REINGRESOS
    ElseIf mi_EstadoEstablecimiento = sghHISEstados.Reingreso Then
        If mi_EstadoServicio = sghHISEstados.Continuador Then
            ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- Si reingresa al establecimiento, no es un continuador"
            If mb_FaltaDato = False Then
                Me.cmbEstadoFrenteServicio.SetFocus
                mb_FaltaDato = True
            End If
            Me.cmbEstadoFrenteServicio.BackColor = ml_ColorError
            Me.cmbEstadoFrenteEstablecimiento.BackColor = ml_ColorError
        End If
    End If
    '
End If

If IdTipoActividad = sghHISTipoActividad.ActividadMasiva Then
    'EDAD Y TIPO DE EDAD
    If CInt(IIf(Replace(Me.txtEdad.Text, "_", "") = "", 0, Replace(Me.txtEdad.Text, "_", ""))) = 0 Then
        ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- No se dígito la Edad"
        If mb_FaltaDato = False Then
            Me.txtEdad.SetFocus
            mb_FaltaDato = True
        End If
        Me.txtEdad.BackColor = ml_ColorError
    Else
        If CInt(Val(mo_cmbTipoEdad.BoundText)) = sghHISTipoEdades.Dias And CInt(IIf(Replace(Me.txtEdad.Text, "_", "") = "", 0, Replace(Me.txtEdad.Text, "_", ""))) > 30 Then
            ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- La edad maxima en días es de 30"
            If mb_FaltaDato = False Then
                Me.txtEdad.SetFocus
                mb_FaltaDato = True
            End If
            Me.txtEdad.BackColor = ml_ColorError
            Me.cmbTipoEdad.BackColor = ml_ColorError
    
        ElseIf CInt(Val(mo_cmbTipoEdad.BoundText)) = sghHISTipoEdades.Meses And CInt(IIf(Replace(Me.txtEdad.Text, "_", "") = "", 0, Replace(Me.txtEdad.Text, "_", ""))) > 11 Then
           
            ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- La edad maxima en meses es de 11"
            If mb_FaltaDato = False Then
                Me.txtEdad.SetFocus
                mb_FaltaDato = True
            End If
            Me.txtEdad.BackColor = ml_ColorError
            Me.cmbTipoEdad.BackColor = ml_ColorError
            
        ElseIf CInt(Val(mo_cmbTipoEdad.BoundText)) = sghHISTipoEdades.Años And CInt(IIf(Replace(Me.txtEdad.Text, "_", "") = "", 0, Replace(Me.txtEdad.Text, "_", ""))) > 99 Then
            ms_ValidacionAtencion = ms_ValidacionAtencion + vbCrLf + "- La edad maxima en años es de 99"
            If mb_FaltaDato = False Then
                Me.txtEdad.SetFocus
                mb_FaltaDato = True
            End If
            Me.txtEdad.BackColor = ml_ColorError
            Me.cmbTipoEdad.BackColor = ml_ColorError
        End If
    End If
End If

ValidarValoresAtencion = ms_ValidacionAtencion
End Function

Function ValidacionDiagnostico(oDODiagnostico As DODiagnostico) As String
Dim mi_genero As Integer
Dim mi_edad As Integer
Dim miTipoEdad As String
Dim ms_MensajeErrorDiagnostico As String

mi_genero = Val(mo_cmbSexo.BoundText)
mi_edad = Val(txtEdad.Text)
miTipoEdad = Val(mo_cmbTipoEdad.BoundText)
ms_MensajeErrorDiagnostico = ""

'Indica que si tiene restriccciones, se validaran con los demas datos
If oDODiagnostico.Restriccion Then
    'valida el diagnostico para el genero
    If oDODiagnostico.IdTipoSexo <> 0 Then
        If mi_genero <> oDODiagnostico.IdTipoSexo Then
            ms_MensajeErrorDiagnostico = "El Diagnóstico no es Correcto para el Género."
        Else
            'valida si el diagnostico es para gestantes
            If Not (oDODiagnostico.Gestacion = True And mi_genero = 2) Then
                ms_MensajeErrorDiagnostico = ms_MensajeErrorDiagnostico & "El Diagnóstico es solo para las Gestantes"
            End If
        End If
    End If
    
    'valida si el diagnostico es intrahospitalario
    If oDODiagnostico.Intrahospitalario Then
        If oDODiagnostico.Intrahospitalario = True Then
            ms_MensajeErrorDiagnostico = ms_MensajeErrorDiagnostico & vbCrLf & "El Diagnóstico es IntraHospitalario."
        End If
    End If
    
    Select Case miTipoEdad
    Case 3 'Anios
        mi_edad = (mi_edad * 30) * 12
    Case 2 'meses
        mi_edad = mi_edad * 30
    Case 1 'dias
        'ya esta convertido
    End Select
    
    If oDODiagnostico.EdadMaxDias < mi_edad Then
        ms_MensajeErrorDiagnostico = ms_MensajeErrorDiagnostico & vbCrLf & "La Edad supera lo permisible para el Diagnóstico."
    End If
    If oDODiagnostico.EdadMinDias > mi_edad Then
        ms_MensajeErrorDiagnostico = ms_MensajeErrorDiagnostico & vbCrLf & "La Edad es inferior a lo permisible para el Diagnóstico."
    End If
    
    'validacion de morbilidad - FALTA
    
End If
ValidacionDiagnostico = Trim(ms_MensajeErrorDiagnostico)
End Function

'CONSULTA DE ATENCIONES INGRESADAS EN EL MOMENTO
Private Sub RefrescarListaAtenciones()
Dim orsTemp As New Recordset
Set orsTemp = mo_ReglasHIS.ObtenerDatosDetalleAtencion(oCabeceraAtencion.IdHisCabecera)

If orsTemp.RecordCount = 0 Then 'Si no hay registros deshabilitamos la opcion de agregar una siguiente hoja
    btnAgregarHoja.Enabled = False
Else
    btnAgregarHoja.Enabled = True
End If

'Limpiamos recordset DetalleAtencion
If oRcs_DetalleAtencion.RecordCount > 0 Then
    oRcs_DetalleAtencion.MoveFirst
    While Not oRcs_DetalleAtencion.EOF
        oRcs_DetalleAtencion.Delete
        oRcs_DetalleAtencion.MoveNext
    Wend
End If

If orsTemp.RecordCount > 0 Then
    orsTemp.MoveFirst
    While Not orsTemp.EOF
        oRcs_DetalleAtencion.AddNew
        oRcs_DetalleAtencion.Fields!IdHisCabecera = orsTemp.Fields!IdHisCabecera
        oRcs_DetalleAtencion.Fields!IdHisDetalle = orsTemp.Fields!IdHisDetalle
        oRcs_DetalleAtencion.Fields!IdTipoAtencion = orsTemp.Fields!IdTipoAtencion
        oRcs_DetalleAtencion.Fields!NroRegistroLote = orsTemp.Fields!NroRegistroLote
        oRcs_DetalleAtencion.Fields!NroRegistroHoja = orsTemp.Fields!NroRegistroHoja
        oRcs_DetalleAtencion.Fields!DiaAtencion = orsTemp.Fields!DiaAtencion
        oRcs_DetalleAtencion.Fields!IdHisPaciente = orsTemp.Fields!IdHisPaciente
        oRcs_DetalleAtencion.Fields!IdPacienteGalenHos = orsTemp.Fields!IdPacienteGalenHos
        oRcs_DetalleAtencion.Fields!idnacionalidad = orsTemp.Fields!idnacionalidad
        oRcs_DetalleAtencion.Fields!IdTipoDocIdentidad = orsTemp.Fields!IdTipoDocIdentidad
        oRcs_DetalleAtencion.Fields!NroDocIdentidad = orsTemp.Fields!NroDocIdentidad
        oRcs_DetalleAtencion.Fields!NroHijo = orsTemp.Fields!NroHijo
        oRcs_DetalleAtencion.Fields!HC_FF_COD = orsTemp.Fields!HC_FF_COD
        oRcs_DetalleAtencion.Fields!IdFinanciador = orsTemp.Fields!IdFinanciador
        oRcs_DetalleAtencion.Fields!IdDistrito = orsTemp.Fields!IdDistrito
        oRcs_DetalleAtencion.Fields!IdEtnia = orsTemp.Fields!IdEtnia
        oRcs_DetalleAtencion.Fields!TipoEdad = orsTemp.Fields!TipoEdad
        oRcs_DetalleAtencion.Fields!Edad = orsTemp.Fields!Edad
        oRcs_DetalleAtencion.Fields!Sexo = orsTemp.Fields!Sexo
        oRcs_DetalleAtencion.Fields!Talla = orsTemp.Fields!Talla
        oRcs_DetalleAtencion.Fields!Peso = orsTemp.Fields!Peso
        oRcs_DetalleAtencion.Fields!IdEstadoaEstablec = orsTemp.Fields!IdEstadoaEstablec
        oRcs_DetalleAtencion.Fields!IdEstadoaServicio = orsTemp.Fields!IdEstadoaServicio
        oRcs_DetalleAtencion.Fields!IdEstado = orsTemp.Fields!IdEstado
        oRcs_DetalleAtencion.Update
        orsTemp.MoveNext
    Wend
    oRcs_DetalleAtencion.MoveFirst
End If
'Set Me.ugvResumenHIS.DataSource = mo_ReglasHIS.ObtenerDatosDetalleAtencion(oCabeceraAtencion.IdHisCabecera)
End Sub

'Funcion que corrige el formato de edad en el control
Private Function DevuelveFormatoEdad(ms_edad As String) As String
ms_edad = Trim(ms_edad)
If Len(ms_edad) = 1 Then
    DevuelveFormatoEdad = "_" & ms_edad
Else
    DevuelveFormatoEdad = ms_edad
End If

End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'AdministrarKeyPreview KeyCode
End Sub
