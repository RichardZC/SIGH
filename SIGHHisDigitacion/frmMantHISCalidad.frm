VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmMantHISCalidad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar hoja HIS"
   ClientHeight    =   7395
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
   Icon            =   "frmMantHISCalidad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
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
      TabIndex        =   20
      Top             =   1560
      Width           =   12855
      Begin VB.TextBox txtNroRegistro 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   120
         MaxLength       =   2
         TabIndex        =   62
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtDia 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   27
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
         TabIndex        =   39
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
         TabIndex        =   37
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmdIngresarDiagnosticos 
         Caption         =   "+Dx"
         Height          =   915
         Left            =   12120
         TabIndex        =   50
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
         TabIndex        =   33
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
         TabIndex        =   30
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   38
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
         TabIndex        =   32
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   31
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
         TabIndex        =   29
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
         TabIndex        =   28
         Top             =   480
         Width           =   1455
      End
      Begin UltraGrid.SSUltraGrid ugvDetalleDiagnosticos 
         Height          =   2655
         Left            =   120
         TabIndex        =   66
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
         TabIndex        =   36
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
      Begin MSMask.MaskEdBox txtTalla 
         Height          =   315
         Left            =   7320
         TabIndex        =   40
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
         TabIndex        =   63
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
         TabIndex        =   57
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
         TabIndex        =   56
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "Tipo Doc."
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
         TabIndex        =   52
         Top             =   240
         Width           =   735
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
         TabIndex        =   21
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
         TabIndex        =   49
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
         TabIndex        =   48
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   43
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
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Nro Doc."
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
         TabIndex        =   24
         Top             =   240
         Width           =   855
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
         TabIndex        =   53
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
         TabIndex        =   23
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
         TabIndex        =   22
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
      TabIndex        =   55
      Top             =   6480
      Width           =   12855
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "frmMantHISCalidad.frx":000C
         DownPicture     =   "frmMantHISCalidad.frx":046C
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
         Picture         =   "frmMantHISCalidad.frx":08E1
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   120
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "frmMantHISCalidad.frx":0D56
         DownPicture     =   "frmMantHISCalidad.frx":121A
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
         Picture         =   "frmMantHISCalidad.frx":1706
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   120
         Width           =   1365
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
         Left            =   9000
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   480
         Width           =   1575
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
         TabIndex        =   10
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
         TabIndex        =   11
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   1080
         Width           =   1815
      End
      Begin MSMask.MaskEdBox txtFechaAnio 
         Height          =   315
         Left            =   10560
         TabIndex        =   67
         Top             =   480
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   9000
         TabIndex        =   61
         Top             =   240
         Width           =   615
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
         Left            =   10560
         TabIndex        =   9
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
         TabIndex        =   6
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
         TabIndex        =   7
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
         TabIndex        =   8
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
         TabIndex        =   4
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
         TabIndex        =   5
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   840
         Width           =   615
      End
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
      Left            =   4440
      TabIndex        =   65
      Top             =   6120
      Width           =   8295
   End
   Begin VB.Label Label15 
      Caption         =   "(F11) Consultar Listas - (F10) Adiciona Diagnóstico"
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
      Index           =   15
      Left            =   0
      TabIndex        =   64
      Top             =   6120
      Width           =   12855
   End
   Begin VB.Label Label15 
      Caption         =   "(F2) Guardar el registro - (F5) Limpiar Atención - (F6) Elminar Diagnóstico"
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
      TabIndex        =   59
      Top             =   5880
      Width           =   12855
   End
   Begin VB.Label Label15 
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
      TabIndex        =   54
      Top             =   6000
      Width           =   11895
   End
End
Attribute VB_Name = "frmMantHISCalidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de HIS calidad
'        Programado por: Cachay F
'        Fecha: Febrero 2014
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

'---------------------------- variables de manejo de negocio -------------------------------
Dim oDOHIS_Detalle_Verifica As New DOHIS_Detalle_Verifica
Dim oCabeceraAtencion As New DOHIS_Cabecera             'Contiene los datos de la cabecera de atencion
Dim mo_ReglasHIS As New SIGHNegocios.ReglasHISGalenos   'Representa la Capa de Negocios del Modulo HIS GalenHos
'Dim mo_DatosParametro As New SIGHDatos.Parametros       'Representa la fecha y hora del servidor
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim oRcs_DetalleAtencion As New Recordset               'Representa el detalle de las Atencion
Dim oRcs_DetalleAtencionTemp As New Recordset
Dim oRcs_Diagnosticos As New Recordset                  'Representa el detalle de Diagnosticos de la Atencion
Dim oRcs_DiagnosticosTemp As New Recordset              'Representa el detalle de Diagnosticos para una Atencion, solo existe por Atencion.

Dim mr_ReglasHIS As New SIGHNegocios.ReglasHISGalenos
Dim ml_IdLote As Long
Dim ml_IdHisDetalle As Long
Dim ml_IdCabeceraHIS As Long
Dim ml_IdUsuario As Long
Dim mo_lcNombrePc As String
Dim mo_lnIdTablaLISTBARITEMS  As Long
Dim mi_Opcion As sghOpciones
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim ml_IdDepartamentoActual As Long: Dim ms_NombreDepActual As String
Dim ml_IdProvinciaActual As Long: Dim ms_NombreProvActual As String
Dim ml_IdDistritoActual As Long: Dim ms_NombreDistrActual As String
Dim ml_IdEstablecimientoActual As Long: Dim ms_CodigoEstablecimiento As String: Dim ms_NombreEstablecimientoActual As String
Dim ml_IdEstablecimiento As Long
Dim mb_SeleccionoLote As Boolean: Dim mb_SeleccionoHoja As Boolean: Dim mb_SeleccionoMedico As Boolean
Dim mb_PesoTallaHabilitados As Boolean
Dim mo_LoteActual As New DOHIS_Lotes
Dim ml_CodigoResponsableDigitacion As Long: Dim ms_NombreRespDigitacion As String
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
Property Let IdHisDetalle(lValue As Long)
   ml_IdHisDetalle = lValue
End Property
Property Get IdHisDetalle() As Long
   IdHisDetalle = ml_IdHisDetalle
End Property
Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
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

Private Sub cmbServicioCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtResponsable
End Sub

Private Sub cmbTurno_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbServicioCodigo
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
    CargarCombosDetalle
    CargarDatosAlFormulario
    mo_Formulario.HabilitarDeshabilitar Me.txtUbigeoDist, False
    mo_Formulario.HabilitarDeshabilitar Me.txtUbigeoEstablecimiento, False
    mo_Formulario.HabilitarDeshabilitar Me.cmbMes, False
    mo_Formulario.HabilitarDeshabilitar Me.txtFechaAnio, False
    mo_Formulario.HabilitarDeshabilitar Me.txtNroPaginas, False
    mo_Formulario.HabilitarDeshabilitar Me.txtLote, False
    mo_Formulario.HabilitarDeshabilitar Me.txtUltimaPaginaLoteActiva, False
    mo_Formulario.HabilitarDeshabilitar Me.txtResponsable, False
    mo_Formulario.HabilitarDeshabilitar Me.txtCodigoEstadistico, False
    mo_Formulario.HabilitarDeshabilitar cmbTurno, False
    mo_Formulario.HabilitarDeshabilitar cmbServicioCodigo, False
    mo_Formulario.HabilitarDeshabilitar txtNroRegistro, False
    txtNacionalidad.Locked = True
    
    Select Case mi_Opcion
    Case sghAgregar
        Me.Caption = "Doble digitación - Ingresar datos registro"
        mb_PesoTallaHabilitados = False
        mb_PrimerIngresoCabeceraAtencion = True
    Case sghModificar, sghConsultar, sghEliminar
        If mi_Opcion = sghModificar Then
            Me.Caption = "Doble digitación - Modificar datos registro"
        ElseIf mi_Opcion = sghConsultar Then
            Me.Caption = "Doble digitación - Consultar datos registro"
            Me.cmdIngresarDiagnosticos.Enabled = False
            Me.btnAceptar.Enabled = False
            BloquearControlesAtencion
        ElseIf mi_Opcion = sghEliminar Then
            Me.Caption = "Doble digitación - Eliminar datos registro"
            Me.cmdIngresarDiagnosticos.Enabled = False
            BloquearControlesAtencion
        End If
    End Select
    mo_Apariencia.ConfigurarFilasBiColores Me.ugvDetalleDiagnosticos, SIGHEntidades.GrillaConFilasBicolor
    'ASIGNAR LOS VALORES POR DEFECTO DEL REGISTRO DE ATENCION
    If mi_Opcion = sghAgregar Then
        ControlesAtencionPorDefecto
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
        mo_Teclado.RealizarNavegacion KeyCode, txtNroHC_FF_COD
    Case vbKeyF2   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
        GrabaAtencionDiagnosticos
    Case vbKeyF5
        IniciaAtencionNueva
    Case vbKeyEscape
        If MsgBox("Desea salir del registro de la doble digitación", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
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
    Case vbKeyF2   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
        GrabaAtencionDiagnosticos
    Case vbKeyF5
        IniciaAtencionNueva
    Case vbKeyEscape
        If MsgBox("Desea salir del registro de la doble digitación", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
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
    Case vbKeyF2   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
        GrabaAtencionDiagnosticos
    Case vbKeyF5
        IniciaAtencionNueva
    Case vbKeyEscape
        If MsgBox("Desea salir del registro de la doble digitación", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
            btnCancelar_Click
        End If
    End Select
End Sub

Private Sub cmdIngresarDiagnosticos_Click()
    Dim Mensaje As String
    Mensaje = ValidarValoresAtencion
    If Len(Mensaje) = 0 Then
        AdicionDiagnostico
    Else
        Call MsgBox("Existen los siguientes problemas:" + vbCrLf + Mensaje, vbInformation, Me.Caption)
        'Me.txtNroRegistro.SetFocus
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
    Case vbKeyF2   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
        GrabaAtencionDiagnosticos
    Case vbKeyF5
        IniciaAtencionNueva
    Case vbKeyEscape
        If MsgBox("Desea salir del registro de la doble digitación", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
            btnCancelar_Click
        End If
    End Select
End Sub

Private Sub txtNroHC_FF_COD_LostFocus()
    IdTipoActividad = CInt(sghHISTipoActividad.Atencion)
    HabilitarDeshabilitarPorNroHcFFCod
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
        Case vbKeyF2   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
            GrabaAtencionDiagnosticos
        Case vbKeyF5
            IniciaAtencionNueva
        Case vbKeyEscape
            If MsgBox("Desea salir del registro de la doble digitación", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
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
            oBusquedaDistrito.DescripcionDistrito = Mid(Me.txtDistritoProcedencia.Text, 10, Len(Me.txtDistritoProcedencia.Text) - 9)
            oBusquedaDistrito.MostrarFormulario
            If oBusquedaDistrito.BotonPresionado = sghAceptar Then
                If oBusquedaDistrito.IdRegistroSeleccionado <> 0 Then
                    ml_IdDistritoAtencion = oBusquedaDistrito.IdRegistroSeleccionado
                    Me.txtDistritoProcedencia.Text = oBusquedaDistrito.IdRegistroSeleccionado & " - " & oBusquedaDistrito.DescripcionRegistroSeleccionado
                    Me.cmbTipoEdad.SetFocus
                Else
                    ml_IdDistritoAtencion = 0
                    Me.txtDistritoProcedencia.Text = "NO ESCOGIDO"
                End If
            End If
            Set oBusquedaDistrito = Nothing
        Case vbKeyBack
            ml_IdDistritoAtencion = 0
        Case vbKeyReturn
            If Me.txtEdad.Enabled = True Then Me.txtEdad.SetFocus
        Case vbKeyF2   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
            GrabaAtencionDiagnosticos
        Case vbKeyF5
            IniciaAtencionNueva
        Case vbKeyEscape
            If MsgBox("Desea salir del registro de la doble digitación", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
                btnCancelar_Click
            End If
    End Select
End Sub

Private Sub txtNroDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn Or vbKeyTab
            'Solo para DNI
            If Val(mo_cmbTipoDocumento.BoundText) = 1 Then
                Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
                Dim mo_HisGalenhos As New SIGHNegocios.ReglasHISGalenos
                Dim mo_DatosFechas As New SIGHEntidades.FechaHora
                Dim o_RcsDatosPaciente As New Recordset
                
                Set o_RcsDatosPaciente = mo_AdminAdmision.PacientesFiltraPorNroDocumentoYtipo(Me.txtNroDocumento.Text, Val(mo_cmbTipoDocumento.BoundText))
                
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
        Case vbKeyF2   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
            GrabaAtencionDiagnosticos
        Case vbKeyF5
            IniciaAtencionNueva
        Case vbKeyEscape
            If MsgBox("Desea salir del registro de la doble digitación", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
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
        Case vbKeyF2   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
            GrabaAtencionDiagnosticos
        Case vbKeyF5
            IniciaAtencionNueva
        Case vbKeyEscape
            If MsgBox("Desea salir del registro de la doble digitación", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
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
        Case vbKeyReturn Or vbKeyTab
            If ml_IdPacienteGalenHos = 0 Then
                cmbEtnia.SetFocus
            Else
                cmbEtnia.SetFocus
            End If
        Case vbKeyF2   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
            GrabaAtencionDiagnosticos
        Case vbKeyF5
            IniciaAtencionNueva
        Case vbKeyEscape
            If MsgBox("Desea salir del registro de la doble digitación", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
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
        Case vbKeyF2   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
            GrabaAtencionDiagnosticos
        Case vbKeyF5
            IniciaAtencionNueva
        Case vbKeyEscape
            If MsgBox("Desea salir del registro de la doble digitación", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
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
'        Case vbKeyF7    'CANCELA EDICION DE ATENCION
'            CancelaEdicionAtencion
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
            mo_Teclado.RealizarNavegacion KeyCode, cmbSexo
        Case vbKeyF2   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
            GrabaAtencionDiagnosticos
        Case vbKeyF5
            IniciaAtencionNueva
        Case vbKeyEscape
            If MsgBox("Desea salir del registro de la doble digitación", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
                btnCancelar_Click
            End If
    End Select
End Sub

Private Sub txtPeso_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn
        mo_Teclado.RealizarNavegacion KeyCode, txtTalla
    Case vbKeyF2   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
        GrabaAtencionDiagnosticos
    Case vbKeyF5
        IniciaAtencionNueva
    Case vbKeyEscape
        If MsgBox("Desea salir del registro de la doble digitación", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
            btnCancelar_Click
        End If
    End Select
End Sub

Private Sub txtTalla_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn
        mo_Teclado.RealizarNavegacion KeyCode, cmbEstadoFrenteServicio
    Case vbKeyF2   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
        GrabaAtencionDiagnosticos
    Case vbKeyF5
        IniciaAtencionNueva
    Case vbKeyEscape
        If MsgBox("Desea salir del registro de la doble digitación", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
            btnCancelar_Click
        End If
    End Select
End Sub

'cmbSexo
Private Sub cmbSexo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            mo_Teclado.RealizarNavegacion KeyCode, txtPeso
        Case vbKeyF2   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
            GrabaAtencionDiagnosticos
        Case vbKeyF5
            IniciaAtencionNueva
        Case vbKeyEscape
            If MsgBox("Desea salir del registro de la doble digitación", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
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
    Case vbKeyF2   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
        GrabaAtencionDiagnosticos
    Case vbKeyF5
        IniciaAtencionNueva
    Case vbKeyEscape
        If MsgBox("Desea salir del registro de la doble digitación", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
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
    Case vbKeyF2   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
        GrabaAtencionDiagnosticos
    Case vbKeyF5
        IniciaAtencionNueva
    Case vbKeyEscape
        If MsgBox("Desea salir del registro de la doble digitación?", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
            btnCancelar_Click
        End If
    End Select
End Sub

Private Sub cmbEstadoFrenteServicio_LostFocus()
   If cmbEstadoFrenteServicio.Text = "" Then
       mo_Formulario.MarcarComoVacio cmbEstadoFrenteServicio
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
    AdministrarKeyPreview CInt(KeyCode)
End Sub

Private Sub ugvDetalleDiagnosticos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.RowSizingArea = ssRowSizingAreaEntireRow
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    Layout.Override.AllowDelete = ssAllowDeleteNo
    
    With Me.ugvDetalleDiagnosticos.Bands(0)
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

Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub btnAceptar_Click()
    GrabaAtencionDiagnosticos
End Sub

'========================================== METODOS ========================================
'CARGA DE LISTADOS DEL FORMUALRIO DE ATENCIONES
Sub CrearTablasTemp()
   
    'para cargar los datos de una consulta
    With oRcs_DiagnosticosTemp
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
    
    
    With oRcs_Diagnosticos
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
    mb_FaltaGrabarAtencion = True
    Dim oRcs_Temp As New ADODB.Recordset
    'OBTENCION DE DATOS CABECERA
    Set oDOHIS_Detalle_Verifica = mo_ReglasHIS.ObtenerDatosHisDetalleVerif(ml_IdHisDetalle)
    Me.txtNroRegistro.Text = oDOHIS_Detalle_Verifica.NroRegistroHoja
    ml_IdCabeceraHIS = oDOHIS_Detalle_Verifica.IdHisCabecera
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
    'OBTENCION DE DATOS DEL LOTE
    mo_LoteActual.IdHisLote = oCabeceraAtencion.IdHisLote 'DEPENDE DE LA CABECERA
    mo_ReglasHIS.ObtenerDatosLotePorIdLote mo_LoteActual
    mo_cmbMes.BoundText = mo_LoteActual.Mes
    txtFechaAnio.Text = mo_LoteActual.Anio
    txtLote.Text = mo_LoteActual.Lote
    txtNroPaginas.Text = mo_LoteActual.NroHojas
    Me.txtUltimaPaginaLoteActiva.Text = oCabeceraAtencion.NroHojaHis 'DEPENDE DE LA CABECERA
    
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
    'Tipo de Documento
    'SERVICIOS DE ESTABLECIMIENTO ACTUAL
    mo_cmbServicioCodigo.BoundColumn = "IdServicio"
    mo_cmbServicioCodigo.ListField = "Nombre"
    Set mo_cmbServicioCodigo.RowSource = mo_ReglasHIS.ListaServiciosPorEstablecimiento(ml_IdEstablecimientoActual)
    Me.cmbServicioCodigo.ListIndex = 0
    'TIPO DE ACTIVIDAD
    IdTipoActividad = sghHISTipoActividad.Atencion
    'CARGAR DATOS DEL DIGITADOR ACTUAL
    Set oRcs_Temp = mo_ReglasHIS.ObtenerDatosDigitador(ml_IdUsuario)
    ml_CodigoResponsableDigitacion = oRcs_Temp.Fields(2)
    ms_NombreRespDigitacion = oRcs_Temp.Fields(1)
    'INGRESO DE VALORES CONTROLES VISUALES
    txtUbigeoDist.Text = ms_CodigoEstablecimiento
    txtUbigeoEstablecimiento.Text = ms_NombreEstablecimientoActual
    txtCodigoEstadistico.Text = ml_CodigoResponsableDigitacion & " - " & ms_NombreRespDigitacion
    mo_cmbTurno.UbicarItemDeComboBoxPorId cmbTurno, oCabeceraAtencion.IdTurno 'poner a equivalencia de su id
    mo_cmbServicioCodigo.UbicarItemDeComboBoxPorId cmbServicioCodigo, oCabeceraAtencion.IdServicio
    If mi_Opcion = sghOpciones.sghConsultar Or mi_Opcion = sghOpciones.sghModificar Or mi_Opcion = sghOpciones.sghEliminar Then
        BuscarRegistro
    End If
End Sub

Private Function BuscarRegistro() As Boolean
    'DEPENDERA DEL TIPO DE REGISTRO - IdTipoActividad
    IdTipoActividad = oDOHIS_Detalle_Verifica.IdTipoAtencion
    Select Case CInt(IdTipoActividad)
    Case sghHISTipoActividad.Atencion
        txtDia.Text = oDOHIS_Detalle_Verifica.DiaAtencion
        txtNroHC_FF_COD.Text = oDOHIS_Detalle_Verifica.nrohc_ff
        IdCodigoNacionalidad = oDOHIS_Detalle_Verifica.idnacionalidad
        Dim oRcs_Temp As New Recordset
        Set oRcs_Temp = mo_ReglasHIS.ObtenerDatosCodNacPorIdNac(IdCodigoNacionalidad)
        oRcs_Temp.MoveFirst
        txtNacionalidad.Text = CStr(oRcs_Temp!Codigo)
        mo_cmbTipoDocumento.BoundText = oDOHIS_Detalle_Verifica.IdTipoDocumento
        txtNroDocumento.Text = oDOHIS_Detalle_Verifica.NroDocIdentidad
        txtOrdenFamiliar.Text = oDOHIS_Detalle_Verifica.NroHijo
        mo_cmbFinanciador.UbicarItemDeComboBoxPorId cmbFinanciador, oDOHIS_Detalle_Verifica.IdTipoFinanciamiento
        mo_cmbEtnia.UbicarItemDeComboBoxPorId cmbEtnia, oDOHIS_Detalle_Verifica.IdEtnia
        ml_IdDistritoAtencion = oDOHIS_Detalle_Verifica.IdDistrito
        Dim oDistrito As New DODistrito
        oDistrito.IdDistrito = ml_IdDistritoAtencion
        mo_ReglasHIS.ConsultarDistritoPorId oDistrito
        txtDistritoProcedencia.Text = ml_IdDistritoAtencion & " - " & oDistrito.Nombre
        mo_cmbTipoEdad.UbicarItemDeComboBoxPorId cmbTipoEdad, oDOHIS_Detalle_Verifica.IdTipoEdad
        txtEdad.Text = DevuelveFormatoEdad(CStr(oDOHIS_Detalle_Verifica.Edad))
        mo_cmbSexo.UbicarItemDeComboBoxPorId cmbSexo, oDOHIS_Detalle_Verifica.Sexo
        txtPeso.Text = IIf(IsNull(oDOHIS_Detalle_Verifica.Peso), "", oDOHIS_Detalle_Verifica.Peso)
        txtTalla.Text = IIf(IsNull(oDOHIS_Detalle_Verifica.Talla), "", oDOHIS_Detalle_Verifica.Talla)
        mo_cmbEstadoFrenteServicio.UbicarItemDeComboBoxPorId cmbEstadoFrenteServicio, oDOHIS_Detalle_Verifica.IdEstadoaServicio
        mo_cmbEstadoFrenteEstablecimiento.UbicarItemDeComboBoxPorId cmbEstadoFrenteEstablecimiento, oDOHIS_Detalle_Verifica.IdEstadoaEstablec
    Case sghHISTipoActividad.ActividadPreventivaPromocional, sghHISTipoActividad.ActividadConAnimales
        txtDia.Text = oDOHIS_Detalle_Verifica.DiaAtencion
        txtNroHC_FF_COD.Text = oDOHIS_Detalle_Verifica.CodigoActividad
    Case sghHISTipoActividad.ActividadMasiva
        txtDia.Text = oDOHIS_Detalle_Verifica.DiaAtencion
        txtNroHC_FF_COD.Text = oDOHIS_Detalle_Verifica.CodigoActividad
        txtEdad.Text = oDOHIS_Detalle_Verifica.Edad
        mo_cmbTipoEdad.UbicarItemDeComboBoxPorId cmbTipoEdad, oDOHIS_Detalle_Verifica.IdTipoEdad
    End Select

    HabilitarCamposAtencionPorActividad CInt(IdTipoActividad)
    
    'DETALLE DIAGNOSTICOS
    Set oRcs_DiagnosticosTemp.DataSource = mo_ReglasHIS.His_ConsultaDxHisDetalleVerif(oDOHIS_Detalle_Verifica.IdHisDetalle)
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
                .Fields!IdHisDetalle = oRcs_DiagnosticosTemp!IdHisDetalle
                .Fields!IdSubClasificacionDX = oRcs_DiagnosticosTemp!IdSubClasificacionDX
                .Fields!IdCIE = oRcs_DiagnosticosTemp!IdCIE
                .Fields!CodLAB = oRcs_DiagnosticosTemp!CodLAB
                .Fields!DESCRIPCION_CIE = oRcs_DiagnosticosTemp!DESCRIPCION_CIE
                .Fields!MSG_ALERTA = oRcs_DiagnosticosTemp!MSG_ALERTA
                .Fields!IdEstado = 0
                .Update
            End With
            oRcs_DiagnosticosTemp.MoveNext
        Loop
    End If
End Function


Sub CargarCombosCabecera()
    mo_cmbMes.BoundColumn = "IdMes"
    mo_cmbMes.ListField = "NombreMes"
    Set mo_cmbMes.RowSource = mo_ReglasHIS.ListaMeses
    'TURNOS
    mo_cmbTurno.BoundColumn = "IdHisTurno"
    mo_cmbTurno.ListField = "Descripcion"
    Set mo_cmbTurno.RowSource = mo_ReglasHIS.ListaTurnos
    Me.cmbTurno.ListIndex = 0
End Sub

'LLENADO DE LISTADOS DE PARA EL FORMULARIO DE INGRESO DE ATENCIONES
Sub CargarCombosDetalle()
    Dim oRcs_Lista As New Recordset
    Set oRcs_Lista = mo_ReglasHIS.ListaTiposDocumento
    oRcs_Lista.MoveFirst
    mo_cmbTipoDocumento.BoundColumn = "IdDocIdentidad"
    mo_cmbTipoDocumento.ListField = "DescripcionLarga"
    Set mo_cmbTipoDocumento.RowSource = oRcs_Lista
    Set oRcs_Lista = Nothing
    
    'Codigo del Tipo de Financiamiento
    Set oRcs_Lista = mo_ReglasHIS.ListaFuentesFinanciamiento
    oRcs_Lista.MoveFirst
    mo_cmbFinanciador.BoundColumn = "IdCodigoFinancHis"
    mo_cmbFinanciador.ListField = "DescripcionLarga"
    Set mo_cmbFinanciador.RowSource = oRcs_Lista
    Set oRcs_Lista = Nothing
    
    'Codigo del Tipo de Etnias
    Set oRcs_Lista = mo_ReglasHIS.ListaEtnias
    oRcs_Lista.MoveFirst
    mo_cmbEtnia.BoundColumn = "codetni"
    mo_cmbEtnia.ListField = "descripcionlarga"
    Set mo_cmbEtnia.RowSource = oRcs_Lista
    Set oRcs_Lista = Nothing
    
    'Codigo del Tipo de Edades
    Set oRcs_Lista = mo_ReglasHIS.ListaTiposEdad
    oRcs_Lista.MoveFirst
    mo_cmbTipoEdad.BoundColumn = "IdHisTipoEdad"
    mo_cmbTipoEdad.ListField = "Descripcionlarga" '"CodigoEdad"
    Set mo_cmbTipoEdad.RowSource = oRcs_Lista

    Set oRcs_Lista = Nothing
    'Codigo del Tipo de Genero
    Set oRcs_Lista = mo_ReglasHIS.ListaTiposSexo
    oRcs_Lista.MoveFirst
    mo_cmbSexo.BoundColumn = "IdTipoSexo"
    mo_cmbSexo.ListField = "Descripcionlarga"
    Set mo_cmbSexo.RowSource = oRcs_Lista
    Set oRcs_Lista = Nothing
    
    'Codigo del Tipo de Estado frente al servicio y al establecimiento
    Set oRcs_Lista = mo_ReglasHIS.ListaSituacionPaciente
    
    oRcs_Lista.MoveFirst
    mo_cmbEstadoFrenteEstablecimiento.BoundColumn = "IdTipoCondicionPaciente"
    mo_cmbEstadoFrenteEstablecimiento.ListField = "Descripcionlarga"
    Set mo_cmbEstadoFrenteEstablecimiento.RowSource = oRcs_Lista
    'cmbEstadoFrenteEstablecimiento.ListIndex = 1
    mo_cmbEstadoFrenteEstablecimiento.BoundText = "N"
    
    oRcs_Lista.MoveFirst
    mo_cmbEstadoFrenteServicio.BoundColumn = "IdTipoCondicionPaciente"
    mo_cmbEstadoFrenteServicio.ListField = "Descripcionlarga"
    Set mo_cmbEstadoFrenteServicio.RowSource = oRcs_Lista
    'cmbEstadoFrenteServicio.ListIndex = 1
    mo_cmbEstadoFrenteServicio.BoundText = "N"
    
    'Listados de valores para la grilla de diagnosticos
    Set oRcs_Lista = Nothing
    'Codigo de Tipos de Diagnosticos
    Set oRcs_Lista = mo_ReglasHIS.ListaTiposDiagnosticos
    oRcs_Lista.MoveFirst
    Me.ugvDetalleDiagnosticos.ValueLists.Add ("ClasificacionDiagnostico")
    While Not oRcs_Lista.EOF
        Me.ugvDetalleDiagnosticos.ValueLists("ClasificacionDiagnostico").ValueListItems.Add CInt(oRcs_Lista!IdSubClasificacionDX), CStr(oRcs_Lista!DescripcionLarga)
        oRcs_Lista.MoveNext
    Wend
    oRcs_Lista.Close
    Set oRcs_Lista = Nothing
End Sub


'ADMINISTAR DESDE AQUI TODAS LAS ACCIONES DE LA GRILLA DE ATENCIONES Y DIAGNOSTICOS
Sub AdministrarKeyPreview(KeyCode As Integer)
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
        If oRcs_Diagnosticos.RecordCount <> 0 Then
           If Not Me.ugvDetalleDiagnosticos.ActiveCell Is Nothing Then
            If Not IsNull(Me.ugvDetalleDiagnosticos.ActiveCell.Column.DataField) Then
                'PARA EL LISTADO DE CODIGOS LAB
                If Me.ugvDetalleDiagnosticos.ActiveCell.Column.DataField = "CodLAB" Then
                    Dim oForm As New frmDetalleCodigosLAB
                    oForm.MostrarFormulario
                    If oForm.BotonPresionado = sghAceptar Then
                        Me.ugvDetalleDiagnosticos.ActiveRow.Cells("CodLAB").Value = oForm.CodigoLab
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
                                oRcs_Temp.Filter = "IdHisDetalle=" & oDOHIS_Detalle_Verifica.IdHisDetalle
                                'oRcs_Temp.MoveFirst
                                
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
    Case vbKeyF2   'GRABA ATENCION Y SUS DIAGNOSTICOS RESPECTIVOS
        GrabaAtencionDiagnosticos
    Case vbKeyF5
        IniciaAtencionNueva
    Case vbKeyF6    'BORRA EL DIAGNOSTICO
        If oRcs_Diagnosticos.RecordCount > 0 Then
            If MsgBox("Desea eliminar el diagnóstico actual?", vbYesNo Or vbQuestion Or vbDefaultButton1, Me.Caption) = vbYes Then
                EliminarDiagnostico
            End If
        End If
    End Select
End Sub

Sub GrabaAtencionDiagnosticos()
    If btnAceptar.Enabled = False Then Exit Sub
    If mi_Opcion = sghConsultar Then Exit Sub
    Select Case mi_Opcion
        Case sghOpciones.sghAgregar
            If ValidarDatosObligatorios Then
                If ValidarReglas Then
                    If mb_FaltaGrabarAtencion Then
                        If MsgBox("Desea guardar los datos de la doble digitación?", vbYesNo Or vbExclamation Or vbDefaultButton1, Me.Caption) = vbYes Then
                            If AdicionAtencion Then
                                Call MsgBox("Se ingreso correctamente los datos de la doble digitación.", vbInformation, Me.Caption)
                                Me.Hide
                            Else
                                Call MsgBox("No se pudo ingresar el registro de la doble digitación, Verificar Error.", vbExclamation, Me.Caption)
                                Exit Sub
                            End If
                        End If
                    End If
                 End If
            End If
        Case sghOpciones.sghModificar
            If ValidarDatosObligatorios Then
                If ValidarReglas Then
                    If mb_FaltaGrabarAtencion Then
                        If MsgBox("Desea guardar los datos de la doble digitación?", vbYesNo Or vbExclamation Or vbDefaultButton1, Me.Caption) = vbYes Then
                            If ModificarAtencion Then
                                Call MsgBox("Se modificó correctamente los datos de la doble digitación.", vbInformation, Me.Caption)
                                mb_FaltaGrabarAtencion = False
                                Me.Hide
                            Else
                                Call MsgBox("No se pudo modificar el registro de la doble digitación, Verificar Error.", vbExclamation, Me.Caption)
                                Exit Sub
                            End If
                         End If
                    End If
                End If
            End If
        Case sghOpciones.sghEliminar
    End Select
End Sub

Function ValidarDatosObligatorios() As Boolean
    On Error Resume Next
    ValidarDatosObligatorios = False
    'Validar número de registros
    If oRcs_DetalleAtencion.RecordCount >= Val(lcBuscaParametro.SeleccionaFilaParametro(272)) Then
        MsgBox "No puede ingresar más de " & lcBuscaParametro.SeleccionaFilaParametro(272) & " registros de atenciones", vbInformation, Me.Caption
        Exit Function
    End If
    'VALIDACION DE VALORES DE ATENCION
    Dim ms_Mensaje As String
    ms_Mensaje = ValidarValoresAtencion
    If Len(ms_Mensaje) > 0 Then
        Call MsgBox("Se encontrarón los siguientes problemas " & vbCrLf & ms_Mensaje, vbInformation, Me.Caption)
        Exit Function
    End If
    ValidarDatosObligatorios = True
End Function

Function ValidarReglas() As Boolean
    Dim ms_mensajeConsistenciaDiagnosticos As String
    ValidarReglas = False
    If mi_Opcion = sghAgregar Or mi_Opcion = sghModificar Then
            Me.ugvDetalleDiagnosticos.Update
            If oRcs_Diagnosticos.RecordCount <= 0 Then
                If IdTipoActividad = sghHISTipoActividad.Atencion Then
                    Call MsgBox("Ingrese los diagnósticos", vbInformation, Me.Caption)
                    Exit Function
                End If
            Else
                oRcs_Diagnosticos.MoveFirst
                Do While Not oRcs_Diagnosticos.EOF
                    If IdTipoActividad <> sghHISTipoActividad.Atencion Then
                       If oRcs_Diagnosticos.Fields!IdSubClasificacionDX <> 102 Then
                            Call MsgBox("Cuando el codigo de activad sea un APP, AMS, AAA los diagnósticos deben ser definitivos(D)", vbInformation, Me.Caption)
                            Exit Function
                       End If
                    End If
                    oRcs_Diagnosticos.MoveNext
                Loop
                ms_mensajeConsistenciaDiagnosticos = mo_ReglasHIS.ValidaConsistenciaDiagnosticosHis(ml_IdLote, CInt(Me.txtDia.Text), Val(Me.txtUltimaPaginaLoteActiva.Text), Val(Me.txtEdad.Text), Val(mo_cmbTipoEdad.BoundText), Val(mo_cmbSexo.BoundText), Val(txtPeso.Text), Me.txtNroHC_FF_COD.Text, Val(mo_cmbEstadoFrenteEstablecimiento.BoundText), Val(mo_cmbEstadoFrenteServicio.BoundText), oRcs_Diagnosticos)
                If Len(ms_mensajeConsistenciaDiagnosticos) <> 0 Then
                    Call MsgBox(ms_mensajeConsistenciaDiagnosticos, vbInformation, Me.Caption)
                    Exit Function
                End If
            End If
            
    End If
    ValidarReglas = True
End Function

Function ExisteDiferencias() As Boolean
    Dim orsTempDetalle As New Recordset
    Dim oRcs_DiagnosticosTemp As New Recordset
    Set orsTempDetalle = mo_ReglasHIS.HIS_ConsultarRegistroDetalleHis(oDOHIS_Detalle_Verifica.IdHisDetalle)
    orsTempDetalle.MoveFirst
    
    ExisteDiferencias = False
    If IdTipoActividad <> orsTempDetalle!IdTipoAtencion Then
        ExisteDiferencias = True
        Exit Function
    End If
    
    'Analizamos el tipo de entrada de los regsitro dependiendo del tipo de atencion
    Select Case CInt(IdTipoActividad)
    
    Case sghHISTipoActividad.Atencion
        If Trim(Me.txtNroHC_FF_COD.Text) <> Trim(orsTempDetalle!HC_FF_COD) Then
            ExisteDiferencias = True
            Exit Function
        End If
        If ml_IdNacionalidadAtencion <> Val(orsTempDetalle!IdPais) Then
            ExisteDiferencias = True
            Exit Function
        End If
        If Val(mo_cmbTipoDocumento.BoundText) <> Val(orsTempDetalle!IdTipoDocumento) Then
            ExisteDiferencias = True
            Exit Function
        End If
        If Me.txtNroDocumento.Text <> Trim(orsTempDetalle!NroDocIdentidad) Then
            ExisteDiferencias = True
            Exit Function
        End If
        If Trim(Me.txtOrdenFamiliar.Text) <> Trim(IIf(IsNull(orsTempDetalle!NroHijo), "", orsTempDetalle!NroHijo)) Then
            ExisteDiferencias = True
            Exit Function
        End If
        If Val(mo_cmbFinanciador.BoundText) <> Val(orsTempDetalle!IdTipoFinanciamiento) Then
            ExisteDiferencias = True
            Exit Function
        End If
        If Val(mo_cmbEtnia.BoundText) <> Val(orsTempDetalle!IdEtnia) Then
            ExisteDiferencias = True
            Exit Function
        End If
        If ml_IdDistritoAtencion <> Val(orsTempDetalle!IdDistrito) Then
            ExisteDiferencias = True
            Exit Function
        End If
        If Val(Me.txtEdad.Text) <> Val(orsTempDetalle!Edad) Then
            ExisteDiferencias = True
            Exit Function
        End If
        If Val(mo_cmbTipoEdad.BoundText) <> Val(orsTempDetalle!IdTipoEdad) Then
            ExisteDiferencias = True
            Exit Function
        End If
        
        If Val(mo_cmbSexo.BoundText) <> orsTempDetalle!Sexo Then
            ExisteDiferencias = True
            Exit Function
        End If
        If Me.txtPeso.Enabled Then
        If Trim(Me.txtPeso.Text) <> IIf(Trim(orsTempDetalle!Peso) = "0", "", Trim(orsTempDetalle!Peso)) Then
            ExisteDiferencias = True
            Exit Function
        End If
        End If
        If Val(IIf(Replace(Me.txtTalla.Text, "_", "") = "", 0, Replace(Me.txtTalla.Text, "_", ""))) <> orsTempDetalle!Talla Then
            ExisteDiferencias = True
            Exit Function
        End If
        If Val(mo_cmbEstadoFrenteEstablecimiento.BoundText) <> Val(orsTempDetalle!IdEstadoaEstablec) Then
            ExisteDiferencias = True
            Exit Function
        End If
        If Val(mo_cmbEstadoFrenteServicio.BoundText) <> Val(orsTempDetalle!IdEstadoaServicio) Then
            ExisteDiferencias = True
            Exit Function
        End If
    Case sghHISTipoActividad.ActividadMasiva
        If Trim(Me.txtNroHC_FF_COD.Text) <> Trim(orsTempDetalle!HC_FF_COD) Then
            ExisteDiferencias = True
            Exit Function
        End If
        If Val(Me.txtEdad.Text) <> Val(orsTempDetalle!Edad) Then
            ExisteDiferencias = True
            Exit Function
        End If
        If Val(mo_cmbTipoEdad.BoundText) <> Val(orsTempDetalle!IdTipoEdad) Then
            ExisteDiferencias = True
            Exit Function
        End If
    Case sghHISTipoActividad.ActividadPreventivaPromocional, sghHISTipoActividad.ActividadConAnimales
        If Trim(Me.txtNroHC_FF_COD.Text) <> Trim(orsTempDetalle!HC_FF_COD) Then
            ExisteDiferencias = True
            Exit Function
        End If
    End Select
    'Validar diagnosticos
    Set oRcs_DiagnosticosTemp = mo_ReglasHIS.ObtenerDatosDetalleDiagnosticoPorIdDetalle(oDOHIS_Detalle_Verifica.IdHisDetalle)
    If oRcs_Diagnosticos.RecordCount <> oRcs_DiagnosticosTemp.RecordCount Then
        ExisteDiferencias = True
        Exit Function
    Else
        oRcs_Diagnosticos.MoveFirst
        oRcs_DiagnosticosTemp.MoveFirst
        Do While Not (oRcs_Diagnosticos.EOF Or oRcs_DiagnosticosTemp.EOF)
            If oRcs_DiagnosticosTemp.Fields!IdCIE <> oRcs_Diagnosticos.Fields!IdCIE Then
                ExisteDiferencias = True
                Exit Function
            End If
            If IIf(IsNull(oRcs_Diagnosticos.Fields!CodLAB), "", oRcs_Diagnosticos.Fields!CodLAB) <> IIf(IsNull(oRcs_DiagnosticosTemp.Fields!CodLAB), "", oRcs_DiagnosticosTemp.Fields!CodLAB) Then
                ExisteDiferencias = True
                Exit Function
            End If
            If oRcs_Diagnosticos.Fields!IdSubClasificacionDX <> oRcs_DiagnosticosTemp.Fields!IdSubClasificacionDX Then
                ExisteDiferencias = True
                Exit Function
            End If
            oRcs_DiagnosticosTemp.MoveNext
            oRcs_Diagnosticos.MoveNext
        Loop
    End If
End Function

'METODO DE ADICION DE UNA ATENCION CON SUS RESPECTIVOS DIAGNOSTICOS
Private Function AdicionAtencion() As Boolean
    mb_FaltaGrabarAtencion = False
    On Error GoTo AdicionAtencion_Error
    If ExisteDiferencias Then
        oDOHIS_Detalle_Verifica.Coincide = 0
    Else
        oDOHIS_Detalle_Verifica.Coincide = 1
    End If
    CargarDatosAlObjetoDatos
    If mo_ReglasHIS.IngresarHISDobleDigitacion(oDOHIS_Detalle_Verifica, oRcs_Diagnosticos) Then
        AdicionAtencion = True
    Else
        AdicionAtencion = False
    End If
    On Error GoTo 0
    Exit Function
AdicionAtencion_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure AdicionAtencion of Formulario frmFormatoHIS"
End Function

Sub CargarDatosAlObjetoDatos()
    Select Case CInt(IdTipoActividad)
        Case sghHISTipoActividad.Atencion
            oDOHIS_Detalle_Verifica.DiaAtencion = Me.txtDia.Text
            oDOHIS_Detalle_Verifica.IdTipoAtencion = IdTipoActividad
            oDOHIS_Detalle_Verifica.nrohc_ff = Trim(Me.txtNroHC_FF_COD.Text)
            oDOHIS_Detalle_Verifica.idnacionalidad = ml_IdNacionalidadAtencion
            oDOHIS_Detalle_Verifica.IdTipoDocumento = Val(mo_cmbTipoDocumento.BoundText)
            oDOHIS_Detalle_Verifica.NroDocIdentidad = Me.txtNroDocumento.Text
            oDOHIS_Detalle_Verifica.NroHijo = Trim(Me.txtOrdenFamiliar.Text)
            oDOHIS_Detalle_Verifica.IdTipoFinanciamiento = Val(mo_cmbFinanciador.BoundText)
            oDOHIS_Detalle_Verifica.IdEtnia = Val(mo_cmbEtnia.BoundText)
            oDOHIS_Detalle_Verifica.IdDistrito = ml_IdDistritoAtencion
            oDOHIS_Detalle_Verifica.Edad = Val(Me.txtEdad.Text)
            oDOHIS_Detalle_Verifica.IdTipoEdad = Val(mo_cmbTipoEdad.BoundText)
            oDOHIS_Detalle_Verifica.Sexo = Val(mo_cmbSexo.BoundText)
            oDOHIS_Detalle_Verifica.Peso = Trim(Me.txtPeso.Text)
            oDOHIS_Detalle_Verifica.Talla = IIf(Trim(Replace(Me.txtTalla.Text, "_", "")) = "", 0, Val(Trim(Replace(Me.txtTalla.Text, "_", ""))))
            oDOHIS_Detalle_Verifica.IdEstadoaEstablec = Val(mo_cmbEstadoFrenteEstablecimiento.BoundText)
            oDOHIS_Detalle_Verifica.IdEstadoaServicio = Val(mo_cmbEstadoFrenteServicio.BoundText)
            oDOHIS_Detalle_Verifica.Registrado = 1
        Case sghHISTipoActividad.ActividadMasiva
            oDOHIS_Detalle_Verifica.DiaAtencion = Me.txtDia.Text
            oDOHIS_Detalle_Verifica.IdTipoAtencion = IdTipoActividad
            oDOHIS_Detalle_Verifica.IdTipoEdad = Val(mo_cmbTipoEdad.BoundText)
            oDOHIS_Detalle_Verifica.Edad = Val(Me.txtEdad.Text)
            oDOHIS_Detalle_Verifica.CodigoActividad = Trim(Me.txtNroHC_FF_COD.Text)
            oDOHIS_Detalle_Verifica.Registrado = 1
        Case sghHISTipoActividad.ActividadPreventivaPromocional, sghHISTipoActividad.ActividadConAnimales
            oDOHIS_Detalle_Verifica.DiaAtencion = Me.txtDia.Text
            oDOHIS_Detalle_Verifica.IdTipoAtencion = IdTipoActividad
            oDOHIS_Detalle_Verifica.CodigoActividad = Trim(Me.txtNroHC_FF_COD.Text)
            oDOHIS_Detalle_Verifica.Registrado = 1
    End Select
End Sub

Private Sub IniciaAtencionNueva()
    If MsgBox("Desea limpiar la pantalla, para registrar nuevamente?", vbYesNo Or vbExclamation Or vbDefaultButton1, Me.Caption) = vbNo Then Exit Sub

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
    'Habilita o Deshabilita segun NroHCFF
    HabilitarDeshabilitarPorNroHcFFCod
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
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ControlesAtencionPorDefecto of Formulario frmMantHISCalidad"
End Sub

Private Function ModificarAtencion() As Boolean
    mb_FaltaGrabarAtencion = False
    On Error GoTo ModificarAtencion_Error
    
    If ExisteDiferencias Then
        oDOHIS_Detalle_Verifica.Coincide = 0
    Else
        oDOHIS_Detalle_Verifica.Coincide = 1
    End If
    
    CargarDatosAlObjetoDatos
    If mo_ReglasHIS.ActualizaHISDobleDigitacion(oDOHIS_Detalle_Verifica, oRcs_Diagnosticos) Then
        ModificarAtencion = True
    Else
        Call MsgBox("Ocurrió un problema con la actualización de la doble digitación.", vbCritical, "HIS Digitación")
        ModificarAtencion = False
    End If
    
    On Error GoTo 0
    Exit Function
ModificarAtencion_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ModificarAtencion of Formulario frmFormatoHIS"
End Function

Private Sub AdicionDiagnostico()
    On Error GoTo AdicionDiagnostico_Error
    
    If oRcs_Diagnosticos.RecordCount < 6 Then
        Dim lbIngresarDx As Boolean
        Dim lnIdCIE As Long
        Dim lcNombreDx As String
        Dim oBusqueda As New SIGHhisDigitacion.BusquedaProductosHis
        Dim oDoFACTCATALOGOSERVICIOS As New sighcomun.DOHis_FactCatalogoServicios
        oBusqueda.CodigoDx = ""
        lnIdCIE = 0
        lcNombreDx = ""
        lbIngresarDx = False
        oBusqueda.MostrarFormulario
        If oBusqueda.BotonPresionado = sghAceptar Then
            lbIngresarDx = True
            lnIdCIE = oBusqueda.IdRegistroSeleccionado
            lcNombreDx = oBusqueda.descripciondiagcpt
            'VALIDACION DE DIAGNOSTICOS REPETIDOS
            Me.ugvDetalleDiagnosticos.Update
            If oRcs_Diagnosticos.RecordCount <> 0 Then
            'SE CLONARA PARA BUSCAR EL DATO
                oRcs_Diagnosticos.MoveFirst
                Do While Not oRcs_Diagnosticos.EOF
                    If Not IsNull(oRcs_Diagnosticos!IdCIE) Then 'indica si es el primero ingresado
                        If lnIdCIE = CLng(oRcs_Diagnosticos!IdCIE) Then
                            lbIngresarDx = False
                            Call MsgBox("El Producto His ya fue ingresado.", vbExclamation, Me.Caption)
                        End If
                    End If
                    oRcs_Diagnosticos.MoveNext
                Loop
                oRcs_Diagnosticos.MoveFirst
            End If
        End If
        Set oBusqueda = Nothing
        If lbIngresarDx = True Then
            With oRcs_Diagnosticos
                .AddNew
                .Fields!IdCIE = lnIdCIE
                .Fields!IdHisDetalle = IdAtencion
                If IdTipoActividad = sghHISTipoActividad.Atencion Then
                'En caso que sea una atencion
                     .Fields!IdSubClasificacionDX = 101
                Else
                'Para los demas sera definitivo
                     .Fields!IdSubClasificacionDX = 102
                End If
                .Fields!DESCRIPCION_CIE = lcNombreDx
                .Fields!CodLAB = ""
                .Fields!IdEstado = 1
                .Update
            End With
            
'             Set ugvDetalleDiagnosticos.DataSource = oRcs_Diagnosticos
             Me.ugvDetalleDiagnosticos.ActiveRow.Activation = ssActivationAllowEdit
             ugvDetalleDiagnosticos.ActiveRow.Cells("DESCRIPCION_CIE").Selected = True
             ugvDetalleDiagnosticos.PerformAction ssKeyActionActivateCell
            ugvDetalleDiagnosticos.PerformAction ssKeyActionEnterEditMode
        End If
    End If

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

'Select Case CInt(Me.ugvDetalleDiagnosticos.ActiveRow.Cells("IdEstado").Value)
'Case 1
    With oRcs_Diagnosticos
        If Not .EOF And Not .BOF Then
            .Delete
            .Update
        End If
    End With
'Case 0, 2 'DESDE LA BASE DE DATOS = 0 , UNA ACTUALIZACION = 2 - Ocultamiento de fila activa y ingreso de valor IDESTADO = 3
'    Me.ugvDetalleDiagnosticos.ActiveRow.Hidden = True
'    Me.ugvDetalleDiagnosticos.ActiveRow.Cells("IdEstado").Value = 3
'End Select
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

'VALIDACION DE LOS CAMPOS INGRESADOS ANTES DE EDITAR LOS DIAGNOSTICOS
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

HabilitarCamposAtencionPorActividad IdTipoActividad

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
            ms_MensajeErrorDiagnostico = "El Diagnostico no es Correcto para el Genero."
        Else
            'valida si el diagnostico es para gestantes
            If Not (oDODiagnostico.Gestacion = True And mi_genero = 2) Then
                ms_MensajeErrorDiagnostico = ms_MensajeErrorDiagnostico & "El Diagnostico es solo para las Gestantes"
            End If
        End If
    End If
    
    'valida si el diagnostico es intrahospitalario
    If oDODiagnostico.Intrahospitalario Then
        If oDODiagnostico.Intrahospitalario = True Then
            ms_MensajeErrorDiagnostico = ms_MensajeErrorDiagnostico & vbCrLf & "El Diagnostico es IntraHospitalario."
        End If
    End If
    
    'valida si el diagnostico esta en los intervalos de edades permisibles
    Select Case miTipoEdad
    Case 3 'Anios
        mi_edad = (mi_edad * 30) * 12
    Case 2 'meses
        mi_edad = mi_edad * 30
    Case 1 'dias
        'ya esta convertido
    End Select
    
    If oDODiagnostico.EdadMaxDias < mi_edad Then
        ms_MensajeErrorDiagnostico = ms_MensajeErrorDiagnostico & vbCrLf & "La Edad supera lo permisible para el Diagnostico."
    End If
    If oDODiagnostico.EdadMinDias > mi_edad Then
        ms_MensajeErrorDiagnostico = ms_MensajeErrorDiagnostico & vbCrLf & "La Edad es inferior a lo permisible para el Diagnostico."
    End If
    
    'validacion de morbilidad - FALTA
    
End If
ValidacionDiagnostico = Trim(ms_MensajeErrorDiagnostico)
End Function

'CONSULTA DE ATENCIONES INGRESADAS EN EL MOMENTO
Private Sub RefrescarListaAtenciones()
Dim orsTemp As New Recordset
Set orsTemp = mo_ReglasHIS.ObtenerDatosDetalleAtencion(oCabeceraAtencion.IdHisCabecera)

'limpiamos recordset DetalleAtencion
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
