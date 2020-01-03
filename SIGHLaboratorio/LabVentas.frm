VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form LabVentas 
   Caption         =   "Form1"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15180
   Icon            =   "LabVentas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   15180
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox grdProductos 
      Height          =   4095
      Left            =   60
      ScaleHeight     =   4035
      ScaleWidth      =   15015
      TabIndex        =   4
      Top             =   3390
      Width           =   15075
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   60
      TabIndex        =   38
      Top             =   7530
      Width           =   15090
      Begin VB.Frame FraRedondeo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   12030
         TabIndex        =   50
         Top             =   180
         Visible         =   0   'False
         Width           =   2925
         Begin VB.TextBox txtRedondeo 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1590
            MaxLength       =   30
            TabIndex        =   51
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Redondear Total"
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
            TabIndex        =   52
            Top             =   300
            Width           =   1365
         End
      End
      Begin VB.CommandButton btnImprimir 
         Caption         =   "Imprime"
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
         Left            =   120
         Picture         =   "LabVentas.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   225
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "LabVentas.frx":11A3
         DownPicture     =   "LabVentas.frx":1667
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
         Left            =   7703
         Picture         =   "LabVentas.frx":1B53
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "LabVentas.frx":203F
         DownPicture     =   "LabVentas.frx":249F
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
         Left            =   6173
         Picture         =   "LabVentas.frx":2914
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame fraCabecera 
      Caption         =   "Datos de Cabecera"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   30
      TabIndex        =   9
      Top             =   60
      Width           =   15105
      Begin VB.CommandButton cmdBuscaCuentaPorApellidos 
         Caption         =   "..."
         Height          =   315
         Left            =   2970
         TabIndex        =   53
         ToolTipText     =   "Busca Cuenta por Apellidos y Nombres"
         Top             =   1590
         Width           =   315
      End
      Begin VB.TextBox txtPlan 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4140
         TabIndex        =   49
         Top             =   2370
         Width           =   4155
      End
      Begin VB.TextBox txtCaja 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   13050
         TabIndex        =   47
         Top             =   690
         Width           =   1845
      End
      Begin VB.TextBox txtNpreventa 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7050
         TabIndex        =   46
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtDatosDeCuenta 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3360
         TabIndex        =   42
         Top             =   1590
         Width           =   4935
      End
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
         Height          =   315
         Left            =   1470
         MaxLength       =   30
         TabIndex        =   0
         Top             =   1590
         Width           =   1425
      End
      Begin VB.Frame fraTipoVenta 
         Height          =   525
         Left            =   1470
         TabIndex        =   36
         Top             =   1020
         Width           =   6825
         Begin Threed.SSOption optVentas 
            Height          =   255
            Left            =   60
            TabIndex        =   43
            Top             =   180
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   450
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Venta Directa"
            Value           =   -1
         End
         Begin Threed.SSOption optPreventa 
            Height          =   255
            Left            =   5580
            TabIndex        =   44
            Top             =   180
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   450
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "PreVenta"
         End
      End
      Begin VB.TextBox txtCajero 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9900
         TabIndex        =   35
         Top             =   1140
         Width           =   4995
      End
      Begin VB.TextBox txtVendedor 
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
         Left            =   9900
         TabIndex        =   34
         Top             =   1560
         Width           =   4995
      End
      Begin VB.TextBox txtTurno 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9900
         TabIndex        =   33
         Top             =   720
         Width           =   1785
      End
      Begin VB.TextBox txtTipoComprobante 
         Enabled         =   0   'False
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
         Left            =   2850
         TabIndex        =   30
         Top             =   270
         Width           =   1215
      End
      Begin VB.TextBox txtDx 
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
         Left            =   9900
         MaxLength       =   30
         TabIndex        =   7
         ToolTipText     =   "Ingrese el Dx (4 digitos)"
         Top             =   2370
         Width           =   1095
      End
      Begin VB.TextBox txtNombreDx 
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
         Height          =   315
         Left            =   11430
         TabIndex        =   17
         Top             =   2370
         Width           =   3465
      End
      Begin VB.CommandButton cmdBuscaDx 
         Caption         =   "..."
         Height          =   315
         Left            =   11070
         TabIndex        =   16
         Top             =   2370
         Width           =   315
      End
      Begin VB.TextBox txtNhistoria 
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
         Left            =   1470
         MaxLength       =   30
         TabIndex        =   1
         ToolTipText     =   "Ingrese el Nro de Historia Clínica"
         Top             =   1980
         Width           =   1425
      End
      Begin VB.TextBox txtNombrePaciente 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3360
         TabIndex        =   15
         Top             =   1980
         Width           =   4935
      End
      Begin VB.CommandButton btnBuscarPaciente 
         Caption         =   "..."
         Height          =   315
         Left            =   2955
         TabIndex        =   14
         ToolTipText     =   "Busca por Apellidos y Nombres"
         Top             =   1980
         Width           =   315
      End
      Begin VB.ComboBox cmbPrescriptor 
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
         Left            =   9900
         TabIndex        =   6
         Top             =   1950
         Width           =   5010
      End
      Begin VB.TextBox txtHoraRegistro 
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
         Height          =   315
         Left            =   11280
         MaxLength       =   30
         TabIndex        =   13
         Top             =   300
         Width           =   735
      End
      Begin VB.ComboBox cmbTipoReceta 
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
         Left            =   1470
         TabIndex        =   3
         Top             =   2760
         Width           =   6870
      End
      Begin VB.ComboBox cmbTipoFinanciamiento 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1470
         TabIndex        =   2
         Top             =   2370
         Width           =   2640
      End
      Begin VB.ComboBox cmbAlmOrigen 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1470
         TabIndex        =   12
         Top             =   660
         Width           =   6840
      End
      Begin VB.TextBox txtObservaciones 
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
         Left            =   9900
         MaxLength       =   100
         TabIndex        =   8
         Top             =   2760
         Width           =   4995
      End
      Begin VB.TextBox txtEstado 
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
         Height          =   315
         Left            =   13050
         MaxLength       =   30
         TabIndex        =   11
         Top             =   270
         Width           =   1845
      End
      Begin VB.TextBox txtDocumento 
         Enabled         =   0   'False
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
         Left            =   1470
         MaxLength       =   30
         TabIndex        =   10
         Top             =   270
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtFregistro 
         Height          =   315
         Left            =   9900
         TabIndex        =   18
         Top             =   300
         Width           =   1350
         _ExtentX        =   2381
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
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Caja"
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
         Left            =   12660
         TabIndex        =   48
         Top             =   750
         Width           =   330
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "N° PreVenta"
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
         Left            =   5910
         TabIndex        =   45
         Top             =   270
         Width           =   1035
      End
      Begin VB.Label lblNcuenta 
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
         TabIndex        =   41
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Venta"
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
         TabIndex        =   37
         Top             =   1140
         Width           =   1185
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Cajero"
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
         Left            =   9375
         TabIndex        =   32
         Top             =   1230
         Width           =   510
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   9390
         TabIndex        =   31
         Top             =   780
         Width           =   495
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Tip.Financiam."
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
         Top             =   2400
         Width           =   1155
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   120
         TabIndex        =   28
         Top             =   1980
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Prescriptor"
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
         Left            =   9015
         TabIndex        =   27
         Top             =   2040
         Width           =   870
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
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
         Left            =   8715
         TabIndex        =   26
         Top             =   2820
         Width           =   1170
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor"
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
         Left            =   9075
         TabIndex        =   25
         Top             =   1650
         Width           =   810
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Diagnóstico"
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
         Left            =   8955
         TabIndex        =   24
         Top             =   2460
         Width           =   930
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Receta"
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
         TabIndex        =   23
         Top             =   2820
         Width           =   990
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
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
         TabIndex        =   22
         Top             =   720
         Width           =   690
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
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
         Left            =   12450
         TabIndex        =   21
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "N° Documento"
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
         TabIndex        =   20
         Top             =   300
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "F.Registro"
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
         Left            =   9090
         TabIndex        =   19
         Top             =   330
         Width           =   810
      End
   End
End
Attribute VB_Name = "LabVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mo_Teclado As New SIGHComun.Teclado
Dim mo_Formulario As New SIGHComun.Formulario
Dim mi_Opcion As sghOpciones
Dim ml_idUsuario As Long
Dim ml_movNumero As String
Dim ml_IdTipoVentaSeleccionada As Long          '0=VentaDirecta      1=PreVenta
Dim mo_cmbAlmacenOrigen As New SIGHComun.ListaDespleglable
Dim mo_cmbPrescriptor As New SIGHComun.ListaDespleglable
Dim mo_cmbTipoFinanciamiento As New SIGHComun.ListaDespleglable
Dim mo_cmbTipoReceta As New SIGHComun.ListaDespleglable
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim oRsTipoFinanciamiento As New Recordset
Dim ms_MensajeError As String
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminAdmision As New ReglasAdmision
Dim ml_IdTipoComprobante As Long
Dim ml_IdDiagnostico As Long
Dim ml_idPaciente As Long
Dim ml_IdVendedor As Long
Dim ml_IdCajero As Long
Dim oBuscaHistoria As New SIGHDatos.Pacientes
Dim mo_DofarmPreVenta As New DoFarmPreVenta
Dim mo_DoPaciente As New DOPaciente
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lnTotalDocumento As Double
Dim mRs_Productos As New Recordset
Dim mo_DoFarmMovimiento As New SIGHComun.DoFarmMovimiento
Dim mo_DoFarmMovimientoVentas As New SIGHComun.DoFarmMovimientoVentas
Dim ml_idTipoConcepto As Long
Const lcConstanteMovimientoSalida As String = "S"
Const lcConstantePreVenta As String = "P"
Const lcConstanteVentaDirecta As String = "D"
Dim ml_EsOficinaElTipoFinanciamiento As Boolean
Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
Dim lcPosicionDefaultCombo As String
Dim ml_IdFuenteFinanciamiento As Long

Property Let movNumero(lValue As String)
   ml_movNumero = lValue
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Let Opcion(lValue As sghOpciones)
   mi_Opcion = lValue
End Property
Property Let TipoVentaSeleccionada(lValue As Long)
   ml_IdTipoVentaSeleccionada = lValue
End Property

Sub ImprimeDocumento()

End Sub

Private Sub btnAceptar_Click()
Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
            CargaDatosAlObjetosDeDatos
            If AgregarDatos() Then
                If optPreventa.Value Then
                   MsgBox "Se agregó correctamente  la Preventa  N° " + txtNpreventa.Text, vbInformation, Me.Caption
                Else
                   MsgBox "Se agregó correctamente  el Documento  N° " + txtDocumento.Text, vbInformation, Me.Caption
                End If
                ImprimeDocumento
                LimpiarDatos
                'Me.Visible = False
            Else
                MsgBox "No se pudo agregar los datos " + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
                grdProductos.RefrescaSaldos
            End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
            CargaDatosAlObjetosDeDatos
            If ModificarDatos() Then
                If optPreventa.Value Then
                   MsgBox "Se Modificó correctamente  la Preventa  N° " + txtNpreventa.Text, vbInformation, Me.Caption
                Else
                   MsgBox "Se Modificó correctamente  el Documento  N° " + txtDocumento.Text, vbInformation, Me.Caption
                End If
                ImprimeDocumento
                Me.Visible = False
                LimpiarVariablesDeMemoria
            Else
                MsgBox "No se pudo modificar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
                grdProductos.RefrescaSaldos
            End If
       End If
   Case sghEliminar
        If MsgBox("Esta seguro de Anular ?", vbQuestion + vbYesNo, "") = vbYes Then
            CargaDatosAlObjetosDeDatos
            If Anular() Then
                If optPreventa.Value Then
                   MsgBox "Se Anuló correctamente  la Preventa  N° " + txtNpreventa.Text, vbInformation, Me.Caption
                Else
                   MsgBox "Se Anuló correctamente  el Documento  N° " + txtDocumento.Text, vbInformation, Me.Caption
                End If
                Me.Visible = False
                LimpiarVariablesDeMemoria
            Else
                MsgBox "No se pudo eliminar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
            End If
        End If
   End Select
End Sub




Private Sub cmbAlmOrigen_Click()
    grdProductos.IdAlmacen = Val(mo_cmbAlmacenOrigen.BoundText)

End Sub

Private Sub cmbAlmOrigen_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbAlmOrigen

End Sub



Private Sub cmbPrescriptor_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbPrescriptor

End Sub

Private Sub cmbTipoFinanciamiento_Click()
    If mo_cmbTipoFinanciamiento.BoundText = "" Then
       Exit Sub
    End If
    oRsTipoFinanciamiento.MoveFirst
    oRsTipoFinanciamiento.Find "idTipoFinanciamiento=" & mo_cmbTipoFinanciamiento.BoundText
    If Not oRsTipoFinanciamiento.EOF Then
       txtTipoComprobante.Text = oRsTipoFinanciamiento.Fields!dComprobante
       ml_IdTipoComprobante = oRsTipoFinanciamiento.Fields!idCajaTiposComprobante
       ml_EsOficinaElTipoFinanciamiento = IIf(oRsTipoFinanciamiento.Fields!esOficina = True, True, False)
       grdProductos.IdTipoFinanciamiento = oRsTipoFinanciamiento.Fields!IdTipoFinanciamiento
       grdProductos.EsOficinaElTipoFinanciamiento = ml_EsOficinaElTipoFinanciamiento
       Select Case oRsTipoFinanciamiento.Fields!IdTipoFinanciamiento
       Case 1        'contado
           ml_idTipoConcepto = 10
       Case 2        'sis
           ml_idTipoConcepto = 13
       Case 3        'soat
           ml_idTipoConcepto = 14
       Case 4        'Convenios
           ml_idTipoConcepto = 23
       Case 5        'Credito Hospitalario
           ml_idTipoConcepto = 17
       Case 6        'Defensa Nacional
           ml_idTipoConcepto = 22
       Case 9        'Exoneraciones
           ml_idTipoConcepto = 15
       Case 10       'credito personal
           ml_idTipoConcepto = 26
       End Select
    Else
       txtTipoComprobante.Text = ""
       ml_IdTipoComprobante = 0
    End If
End Sub



Private Sub cmbTipoFinanciamiento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbTipoFinanciamiento
  
End Sub



Private Sub cmbTipoReceta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbTipoReceta

End Sub

Private Sub cmdBuscaCuentaPorApellidos_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaPacientes
    Dim oDOPaciente As New DOPaciente
    oBusqueda.TipoFiltro = sghFiltrarConHistoriasTemporales
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
        If Not oDOPaciente Is Nothing Then
            ml_idPaciente = oDOPaciente.idPaciente
            txtNhistoria.Text = oDOPaciente.NroHistoriaClinica
            txtNombrePaciente.Text = Trim(oDOPaciente.ApellidoPaterno) + " " + Trim(oDOPaciente.ApellidoMaterno) + " " + oDOPaciente.PrimerNombre
            Dim oRsTmp As New Recordset
            Set oRsTmp = mo_ReglasFarmacia.FacturacionCuentasAtencionSeleccionarPorIdPaciente(ml_idPaciente)
            If oRsTmp.RecordCount > 0 Then
               txtNcuenta.Text = oRsTmp.Fields!idCuentaAtencion
            End If
            oRsTmp.Close
            Set oRsTmp = Nothing
            txtNcuenta_LostFocus
        End If
    End If
End Sub

Private Sub Form_Activate()
   If mi_Opcion <> sghAgregar Then
      cmbTipoReceta.SetFocus
   End If
End Sub

Private Sub Form_Initialize()
    Set mo_cmbAlmacenOrigen.MiComboBox = cmbAlmOrigen
    Set mo_cmbPrescriptor.MiComboBox = cmbPrescriptor
    Set mo_cmbTipoFinanciamiento.MiComboBox = cmbTipoFinanciamiento
    Set mo_cmbTipoReceta.MiComboBox = cmbTipoReceta
    
End Sub

Private Sub Form_Load()
    ConfigurarGrdProductos
    CargarComboBoxes
    
    Select Case mi_Opcion
    Case sghAgregar
        Me.Caption = "Agregar Ventas"
    Case sghModificar
        Me.Caption = "Modificar Ventas"
    Case sghConsultar
        Me.Caption = "Consultar Ventas"
        btnImprimir.Visible = True
    Case sghEliminar
        Me.Caption = "Anular Ventas"
    End Select
    CargarDatosAlFormulario
End Sub
Sub ConfigurarGrdProductos()
    grdProductos.movNumero = ml_movNumero
    grdProductos.IdAlmacen = 0
    grdProductos.Inicializar
    grdProductos.TipoPrecioParaNiNs = 3    'precio de venta

End Sub

Sub CargarDatosAlFormulario()
    mo_Formulario.HabilitarDeshabilitar Me.txtDocumento, False
    mo_Formulario.HabilitarDeshabilitar Me.txtNpreventa, False
    mo_Formulario.HabilitarDeshabilitar Me.txtTipoComprobante, False
    mo_Formulario.HabilitarDeshabilitar Me.txtFregistro, False
    mo_Formulario.HabilitarDeshabilitar Me.txtHoraRegistro, False
    mo_Formulario.HabilitarDeshabilitar Me.txtEstado, False
    mo_Formulario.HabilitarDeshabilitar Me.txtNombrePaciente, False
    mo_Formulario.HabilitarDeshabilitar Me.txtNombreDx, False
    mo_Formulario.HabilitarDeshabilitar Me.txtTurno, False
    mo_Formulario.HabilitarDeshabilitar Me.txtVendedor, False
    mo_Formulario.HabilitarDeshabilitar Me.txtCajero, False
    mo_Formulario.HabilitarDeshabilitar Me.txtCaja, False
    mo_Formulario.HabilitarDeshabilitar Me.txtDatosDeCuenta, False
    mo_Formulario.HabilitarDeshabilitar Me.txtPlan, False
    Select Case mi_Opcion
     Case sghAgregar
        txtFregistro.Text = lcBuscaParametro.RetornaFechaServidorSQL      'Format(Now, "HH:MM")
        grdProductos.movNumero = ""
        grdProductos.LimpiarGrilla
        grdProductos.CargaProductosPorMovNumero
        grdProductos.AgregaRegistro
        optVentas_Click 1
     Case sghModificar
        DeshabilitaCabecera
        CargarDatosALosControles
     Case sghConsultar
        DeshabilitaCabecera
        CargarDatosALosControles
        btnAceptar.Enabled = False
     Case sghEliminar
        DeshabilitaCabecera
        CargarDatosALosControles
 End Select
End Sub

Sub DeshabilitaCabecera()
    If ml_IdTipoVentaSeleccionada = 0 Then
        mo_Formulario.HabilitarDeshabilitar Me.cmbAlmOrigen, False
        fraTipoVenta.Enabled = False
        mo_Formulario.HabilitarDeshabilitar Me.txtNcuenta, False
        cmdBuscaCuentaPorApellidos.Enabled = False
        mo_Formulario.HabilitarDeshabilitar Me.txtNhistoria, False
        btnBuscarPaciente.Enabled = False
        mo_Formulario.HabilitarDeshabilitar Me.cmbTipoFinanciamiento, False
    Else
        mo_Formulario.HabilitarDeshabilitar Me.cmbAlmOrigen, False
        fraTipoVenta.Enabled = False
        mo_Formulario.HabilitarDeshabilitar Me.txtNcuenta, False
        cmdBuscaCuentaPorApellidos.Enabled = False
        mo_Formulario.HabilitarDeshabilitar Me.txtNhistoria, False
        btnBuscarPaciente.Enabled = False
        mo_Formulario.HabilitarDeshabilitar Me.cmbTipoFinanciamiento, False
    End If
End Sub


Sub CargarComboBoxes()
    Dim lnIdAlmacen As Long
    Dim oBuscaDondeLabora As New SIGHNegocios.ReglasComunes
    lnIdAlmacen = oBuscaDondeLabora.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghAlmacenFarmacia, ml_idUsuario)
    Set oBuscaDondeLabora = Nothing
    mo_cmbAlmacenOrigen.BoundColumn = "IdAlmacen"
    mo_cmbAlmacenOrigen.ListField = "Descripcion"
    Set mo_cmbAlmacenOrigen.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idtipoLocales='F' and idTipoSuministro='01' and idEstado=1")
    ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
    If lnIdAlmacen > 0 Then
       mo_cmbAlmacenOrigen.BoundText = lnIdAlmacen
       mo_Formulario.HabilitarDeshabilitar Me.cmbAlmOrigen, False
    End If
   '
    mo_cmbPrescriptor.BoundColumn = "idEmpleado"
    mo_cmbPrescriptor.ListField = "ApNom"
    Set mo_cmbPrescriptor.RowSource = mo_ReglasFarmacia.EmpleadosDevuelvePrescriptores
    ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
   '
    mo_cmbTipoReceta.BoundColumn = "idTipoReceta"
    mo_cmbTipoReceta.ListField = "TipoReceta"
    Set mo_cmbTipoReceta.RowSource = mo_ReglasFarmacia.FarmTipoRecetaDevuelveTodos
    ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
    '
    Set oRsTipoFinanciamiento = mo_ReglasFarmacia.TipoFinanciamientosDevuelveSoloFarmacia("")
    mo_cmbTipoFinanciamiento.BoundColumn = "idTipoFinanciamiento"
    mo_cmbTipoFinanciamiento.ListField = "Descripcion"
    Set mo_cmbTipoFinanciamiento.RowSource = mo_ReglasFarmacia.TipoFinanciamientosDevuelveSoloFarmacia("")
    ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
    '
    If ms_MensajeError <> "" Then
       MsgBox ms_MensajeError
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub





Private Sub grdProductos_SePresionoTeclaEspecial(KeyCode As Integer)
     If KeyCode = vbKeyF2 Then
        AdministrarKeyPreview KeyCode
        Me.KeyPreview = False
        'SendKeys "{F2}"
     End If
End Sub

Private Sub grdProductos_Totalizado(lnTotalIngresado As Double)
    txtRedondeo.Text = lnTotalIngresado
End Sub

Private Sub optPreventa_Click(Value As Integer)
    If optPreventa.Value Then
        Set oRsTipoFinanciamiento = mo_ReglasFarmacia.TipoFinanciamientosDevuelveSoloFarmacia(" and dbo.TiposFinanciamiento.TipoVenta='P'")
        lcPosicionDefaultCombo = ""
        If oRsTipoFinanciamiento.RecordCount = 1 Then
            lcPosicionDefaultCombo = Trim(Str(oRsTipoFinanciamiento.Fields!IdTipoFinanciamiento))
        End If
        mo_cmbTipoFinanciamiento.BoundColumn = "idTipoFinanciamiento"
        mo_cmbTipoFinanciamiento.ListField = "Descripcion"
        Set mo_cmbTipoFinanciamiento.RowSource = mo_ReglasFarmacia.TipoFinanciamientosDevuelveSoloFarmacia(" and dbo.TiposFinanciamiento.TipoVenta='P'")
        If lcPosicionDefaultCombo <> "" Then
           mo_cmbTipoFinanciamiento.BoundText = lcPosicionDefaultCombo
        End If
        ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
        mo_Formulario.HabilitarDeshabilitar Me.cmbTipoFinanciamiento, True
        mo_Formulario.HabilitarDeshabilitar Me.txtNhistoria, True
        mo_Formulario.HabilitarDeshabilitar Me.txtObservaciones, False
        btnBuscarPaciente.Enabled = True
        If mi_Opcion = sghAgregar Then
           cmbTipoReceta.Text = ""
        End If
        FraRedondeo.Visible = True
    End If
End Sub

Private Sub optVentas_Click(Value As Integer)
    If optVentas.Value Then
        Set oRsTipoFinanciamiento = mo_ReglasFarmacia.TipoFinanciamientosDevuelveSoloFarmacia(" and dbo.TiposFinanciamiento.TipoVenta='D'")
        mo_cmbTipoFinanciamiento.BoundColumn = "idTipoFinanciamiento"
        mo_cmbTipoFinanciamiento.ListField = "Descripcion"
        Set mo_cmbTipoFinanciamiento.RowSource = mo_ReglasFarmacia.TipoFinanciamientosDevuelveSoloFarmacia(" and dbo.TiposFinanciamiento.TipoVenta='D'")
        grdProductos.TipoVentaSeleccionada = 0
        ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
        mo_Formulario.HabilitarDeshabilitar Me.cmbTipoFinanciamiento, False
        mo_Formulario.HabilitarDeshabilitar Me.txtNhistoria, False
        mo_Formulario.HabilitarDeshabilitar Me.txtObservaciones, True
        btnBuscarPaciente.Enabled = False
        If mi_Opcion = sghAgregar Then
           mo_cmbTipoReceta.BoundText = "1"
        End If
        FraRedondeo.Visible = False
    End If
End Sub

Private Sub txtDx_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDx

End Sub


Private Sub txtDx_LostFocus()
        Dim oDODiagnostico As DODiagnostico
        Set oDODiagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorCodigoCIE2004(txtDx.Text)
        If Not oDODiagnostico Is Nothing Then
            ml_IdDiagnostico = oDODiagnostico.idDiagnostico
            txtNombreDx.Text = oDODiagnostico.descripcion
        Else
            ml_IdDiagnostico = 0
            txtNombreDx.Text = ""
        End If

End Sub

Private Sub txtNcuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNcuenta

End Sub


Private Sub txtNcuenta_LostFocus()
   If Val(txtNcuenta.Text) > 0 Then
       Dim oRsTmp As New Recordset
       Dim lbSigue As Boolean
       Set oRsTmp = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(txtNcuenta.Text)
       txtDatosDeCuenta.Text = ""
       If mi_Opcion = sghAgregar Then
          cmbTipoFinanciamiento.Text = ""
          txtTipoComprobante.Text = ""
       End If
       ml_idPaciente = 0
       ml_IdFuenteFinanciamiento = 0
       txtNombrePaciente.Text = ""
       txtNhistoria.Text = ""
       
       txtPlan.Text = ""
       lbSigue = True
       If oRsTmp.RecordCount > 0 Then
          If oRsTmp.Fields!idEstado <> 1 Then
             If mi_Opcion <> sghConsultar Then
                MsgBox "Ese estado de Cuenta no se encuentra ABIERTA", vbInformation, Me.Caption
                If mi_Opcion = sghModificar Or mi_Opcion = sghEliminar Then
                   btnAceptar.Enabled = False
                Else
                   lbSigue = False
                End If
             End If
          End If
          If lbSigue Then
                txtDatosDeCuenta.Text = "F.Ing: " & oRsTmp.Fields!fechaIngreso & " - " & IIf(oRsTmp.Fields!IdTipoServicio = 1, "Consultorios Externos", IIf(oRsTmp.Fields!IdTipoServicio = 3, "Hospitalización", "Emergencia")) & " - (Est: " & Trim(oRsTmp.Fields!estadoCta) & ")"
                txtPlan.Text = "Plan Act.: " & oRsTmp.Fields!dFuenteFinanciamiento
                ml_idPaciente = oRsTmp.Fields!idPaciente
                ml_IdFuenteFinanciamiento = oRsTmp.Fields!idFuenteFinanciamiento
                txtNombrePaciente.Text = Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & " " & oRsTmp.Fields!PrimerNombre
                txtNhistoria.Text = oRsTmp.Fields!NroHistoriaClinica
                If mi_Opcion = sghAgregar Then
                      If optVentas.Value = True Then
                         mo_cmbTipoFinanciamiento.BoundText = oRsTmp.Fields!idFormaPago
                      Else
                         cmbTipoFinanciamiento_Click
                      End If
                      
'                      If oRsTmp.Fields!tipoVenta = "D" Then
'                          optVentas.Value = True
'                          optPreventa.Value = False
'                          optVentas_Click (1)
'                      Else
'                          optVentas.Value = False
'                          optPreventa.Value = True
'                          optPreventa_Click (1)
'                      End If
'                      mo_cmbTipoFinanciamiento.BoundText = oRsTmp.Fields!idFormaPago
'                      cmbTipoFinanciamiento_Click
                End If
          End If
       End If
       oRsTmp.Close
       Set oRsTmp = Nothing
   End If
End Sub

Private Sub txtNhistoria_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNhistoria

End Sub


Private Sub btnBuscarPaciente_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaPacientes
    Dim oDOPaciente As New DOPaciente
    oBusqueda.TipoFiltro = sghFiltrarConHistoriasTemporales
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
        If Not oDOPaciente Is Nothing Then
            ml_idPaciente = oDOPaciente.idPaciente
            txtNhistoria.Text = oDOPaciente.NroHistoriaClinica
            txtNombrePaciente.Text = Trim(oDOPaciente.ApellidoPaterno) + " " + Trim(oDOPaciente.ApellidoMaterno) + " " + oDOPaciente.PrimerNombre
        End If
    End If
End Sub
Private Sub cmdBuscaDx_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaDiagnosticos
    Dim oDODiagnostico As DODiagnostico
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDODiagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
        If Not oDODiagnostico Is Nothing Then
            ml_IdDiagnostico = oDODiagnostico.idDiagnostico
            txtDx.Text = oDODiagnostico.CodigoCIE2004
            txtNombreDx.Text = oDODiagnostico.descripcion
        End If
    End If
End Sub


Sub CargaDatosVendedorCajero(lnIdVendedorCajero As Long, EsVendedor As Boolean)
    Dim oDOEmpleado As dOEmpleado
    Set oDOEmpleado = mo_AdminServiciosComunes.EmpleadosSeleccionarPorId(lnIdVendedorCajero)
    With oDOEmpleado
        If EsVendedor Then
           txtVendedor.Text = Trim(.ApellidoPaterno) & " " & Trim(.ApellidoMaterno) & .nombres
        Else
           txtCajero.Text = Trim(.ApellidoPaterno) & " " & Trim(.ApellidoMaterno) & .nombres
        End If
    End With
End Sub

Sub CargarDatosALosControles()
   If ml_IdTipoVentaSeleccionada = 0 Then
        CargaVentasDirectas
        '******permiso a Modificar documento con Fecha Anterior a la actual
        Dim mo_PermisosFacturacion As New PermisosFacturacion
        Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
        Set mo_PermisosFacturacion = mo_ReglasSeguridad.UsuariosRolesSeleccionarPermisosFacturacion(ml_idUsuario)
        If mo_PermisosFacturacion.ActualizaFechaDocumentoES = False Then
           If CDate(lcBuscaParametro.RetornaFechaServidorSQL) <> CDate(txtFregistro.Text) Then
              MsgBox "No tiene ACCESO a Modificar/Anular una Venta" & Chr(13) & " de una Fecha Registro diferente a la actual", vbExclamation, Me.Caption
              btnAceptar.Enabled = False
           End If
        End If
        Set mo_PermisosFacturacion = Nothing
        Set mo_ReglasSeguridad = Nothing
   Else
        CargaPreVenta
   End If
   DeshabilitaCabecera
  
End Sub
Sub CargaPreVenta()
    mo_DofarmPreVenta.idPreVenta = Val(ml_movNumero)
    If Not mo_ReglasFarmacia.FarmPreventasSeleccionarPorId(mo_DofarmPreVenta) Then
      MsgBox mo_ReglasFarmacia.MensajeError
      Exit Sub
    End If
    txtNpreventa.Text = ml_movNumero
    mo_cmbAlmacenOrigen.BoundText = mo_DofarmPreVenta.IdAlmacen
    txtNcuenta.Text = IIf(mo_DofarmPreVenta.idCuentaAtencion > 0, mo_DofarmPreVenta.idCuentaAtencion, "")
    txtNcuenta_LostFocus
    optVentas.Value = False: optPreventa.Value = True
    mo_cmbTipoFinanciamiento.BoundText = mo_DofarmPreVenta.IdTipoFinanciamiento
    mo_cmbTipoReceta.BoundText = mo_DofarmPreVenta.idTipoReceta
    ml_IdVendedor = mo_DofarmPreVenta.idVendedor
    CargaDatosVendedorCajero ml_IdVendedor, True
    mo_cmbPrescriptor.BoundText = mo_DofarmPreVenta.idPrescriptor
    'Paciente
    If txtNcuenta.Text = "" Then
        ml_idPaciente = mo_DofarmPreVenta.idPaciente
        If ml_idPaciente > 0 Then
            mo_DoPaciente.idPaciente = ml_idPaciente
            Set mo_DoPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(ml_idPaciente)
            txtNhistoria.Text = mo_DoPaciente.NroHistoriaClinica
            txtNombrePaciente.Text = Trim(mo_DoPaciente.ApellidoPaterno) & " " & Trim(mo_DoPaciente.ApellidoMaterno) & " " & mo_DoPaciente.PrimerNombre
        End If
    End If
    'Dx
    Dim mo_Diagnostico As New DODiagnostico
    ml_IdDiagnostico = mo_DofarmPreVenta.idDiagnostico
    If ml_IdDiagnostico > 0 Then
        Set mo_Diagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorId(ml_IdDiagnostico)
        txtDx.Text = mo_Diagnostico.CodigoCIE2004
        txtNombreDx.Text = mo_Diagnostico.descripcion
    End If
   '**************Datos de la tabla FarmMovimientoDetalle *****************
   grdProductos.idPreVenta = Val(ml_movNumero)
   grdProductos.TipoVentaSeleccionada = 1   'Preventas
   grdProductos.CargaProductosPorIdPreVenta
   grdProductos.RefrescarDatos
   lnTotalDocumento = grdProductos.DevuelveTotal
   txtEstado.Text = mo_ReglasFarmacia.DevuelveEstadoDocumento("select estado as lcEstado from FarmEstadosPreVenta where idEstadoPreventa=" & mo_DofarmPreVenta.idEstadoPreventa)
   If mo_DofarmPreVenta.idEstadoPreventa <> 1 Then
      btnAceptar.Enabled = False
   End If
   If mo_DofarmPreVenta.idEstadoPreventa = 2 Or mo_DofarmPreVenta.idEstadoPreventa = 0 Then
      Dim oRsTmp As New Recordset
      Dim oDoComprobantePago As New DOCajaComprobantesPago
      Set oRsTmp = mo_ReglasFarmacia.FarmMovimientoVentasSeleccionarPorIdPreventa(mo_DofarmPreVenta.idPreVenta)
      If oRsTmp.RecordCount > 0 Then
         txtDocumento.Text = oRsTmp.Fields!DocumentoNumero
      End If
      oRsTmp.Close
      Set oRsTmp = mo_ReglasCaja.CajaComprobantePagoSeleccionarPorFechaOdocumento(Left(txtDocumento.Text, 3), Mid(txtDocumento.Text, 5, 15), Date, Date)
      If oRsTmp.RecordCount > 0 Then
         txtCaja.Text = IIf(IsNull(oRsTmp.Fields!dCaja), "", oRsTmp.Fields!dCaja)
         txtCajero.Text = IIf(IsNull(oRsTmp.Fields!ApellidoPaterno), "", Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & " " & oRsTmp.Fields!nombres)
         txtTurno.Text = IIf(IsNull(oRsTmp.Fields!descripcion), "", oRsTmp.Fields!descripcion)
      End If
      oRsTmp.Close
      Set oRsTmp = Nothing
   End If
   If lnTotalDocumento <> mo_DofarmPreVenta.Total Then
      FraRedondeo.Visible = True
      txtRedondeo.Text = mo_DofarmPreVenta.Total
   End If
End Sub
Sub CargaVentasDirectas()
 '**************Datos de la tabla FarmMovimiento *****************
   mo_DoFarmMovimiento.movNumero = ml_movNumero
   mo_DoFarmMovimiento.MovTipo = lcConstanteMovimientoSalida
   If Not mo_ReglasFarmacia.FarmMovimientoSeleccionarPorId(mo_DoFarmMovimiento) Then
      MsgBox mo_ReglasFarmacia.MensajeError
      Exit Sub
   End If
   txtDocumento.Text = mo_DoFarmMovimiento.DocumentoNumero
   mo_cmbAlmacenOrigen.BoundText = mo_DoFarmMovimiento.IdAlmacenOrigen
   txtObservaciones.Text = mo_DoFarmMovimiento.Observaciones
   txtEstado.Text = mo_ReglasFarmacia.DevuelveEstadoDocumento("select estado as lcEstado from FarmEstadosMovimientos where idEstadoMovimiento=" & mo_DoFarmMovimiento.idEstadoMovimiento)
   txtFregistro.Text = Format(mo_DoFarmMovimiento.FechaCreacion, "dd/mm/yyyy")
   txtHoraRegistro.Text = mo_DoFarmMovimiento.HoraCreacion

   '**************Datos de la tabla FarmMovimientoVentas *****************
   
   Dim mo_DoPaciente As New DOPaciente
   Dim mo_Diagnostico As New DODiagnostico
   Dim oConexion As New ADODB.Connection
   With mo_DoFarmMovimientoVentas
       .movNumero = ml_movNumero
       .MovTipo = lcConstanteMovimientoSalida
       If Not mo_ReglasFarmacia.FarmMovimientoVentasSeleccionarPorId(mo_DoFarmMovimientoVentas) Then
            MsgBox mo_ReglasFarmacia.MensajeError
            Exit Sub
       Else
            txtNpreventa.Text = .idPreVenta
            mo_cmbPrescriptor.BoundText = .idPrescriptor
            txtNcuenta.Text = .idCuentaAtencion
            txtNcuenta_LostFocus
            optPreventa.Value = False: optVentas.Value = True
            mo_cmbTipoFinanciamiento.BoundText = .IdTipoFinanciamiento
            mo_cmbTipoReceta.BoundText = .idTipoReceta
            mo_cmbPrescriptor.BoundText = .idPrescriptor
            'Dx
            ml_IdDiagnostico = .idDiagnostico
            If ml_IdDiagnostico > 0 Then
                Set mo_Diagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorId(.idDiagnostico)
                txtDx.Text = mo_Diagnostico.CodigoCIE2004
                txtNombreDx.Text = mo_Diagnostico.descripcion
                Set mo_Diagnostico = Nothing
            End If
       End If
   End With
   If mo_cmbTipoFinanciamiento.BoundText = "1" And txtNpreventa.Text <> "" Then
      Dim oRsTmp As New Recordset
      Dim oDoComprobantePago As New DOCajaComprobantesPago
      Set oRsTmp = mo_ReglasCaja.CajaComprobantePagoSeleccionarPorFechaOdocumento(Left(txtDocumento.Text, 3), Mid(txtDocumento.Text, 5, 15), Date, Date)
      If oRsTmp.RecordCount > 0 Then
         txtCaja.Text = IIf(IsNull(oRsTmp.Fields!dCaja), "", oRsTmp.Fields!dCaja)
         txtCajero.Text = IIf(IsNull(oRsTmp.Fields!ApellidoPaterno), "", Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & " " & oRsTmp.Fields!nombres)
         txtTurno.Text = IIf(IsNull(oRsTmp.Fields!descripcion), "", oRsTmp.Fields!descripcion)
      End If
      oRsTmp.Close
      Set oRsTmp = Nothing
      mo_DofarmPreVenta.idPreVenta = Val(txtNpreventa.Text)
      If Not mo_ReglasFarmacia.FarmPreventasSeleccionarPorId(mo_DofarmPreVenta) Then
         MsgBox mo_ReglasFarmacia.MensajeError
         Exit Sub
      End If
      If Val(txtNcuenta.Text) = 0 And mo_DofarmPreVenta.idPaciente > 0 Then
         mo_DoPaciente.ApellidoPaterno = ""
         Set mo_DoPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(mo_DofarmPreVenta.idPaciente)
         If mo_DoPaciente.ApellidoPaterno <> "" Then
             ml_idPaciente = mo_DofarmPreVenta.idPaciente
             txtNombrePaciente.Text = Trim(mo_DoPaciente.ApellidoPaterno) & " " & Trim(mo_DoPaciente.ApellidoMaterno) & " " & mo_DoPaciente.PrimerNombre
             txtNhistoria.Text = mo_DoPaciente.NroHistoriaClinica
         End If
      End If
      CargaDatosVendedorCajero mo_DofarmPreVenta.idVendedor, True
   End If
   '**************Datos de la tabla FarmMovimientoDetalle *****************
   grdProductos.movNumero = ml_movNumero
   grdProductos.TipoVentaSeleccionada = 0   'VentaDirecta
   grdProductos.CargaProductosPorMovNumero
   grdProductos.RefrescarDatos
   lnTotalDocumento = grdProductos.DevuelveTotal
   'If mo_DoFarmMovimiento.IdEstadoMovimiento <> 1 Then
   '   btnAceptar.Enabled = False
   'End If
   If Val(txtNpreventa.Text) > 0 Then
      btnAceptar.Enabled = False
   End If
   If mo_cmbTipoFinanciamiento.BoundText = "1" And txtNpreventa.Text <> "" Then
        If lnTotalDocumento <> mo_DofarmPreVenta.Total Then
           FraRedondeo.Visible = True
           txtRedondeo.Text = mo_DoFarmMovimiento.Total
        End If
   End If
End Sub
Function ValidarDatosObligatorios() As Boolean
   ValidarDatosObligatorios = False
   ms_MensajeError = ""
   If optPreventa.Value Then
        If cmbAlmOrigen.Text = "" Then
            ms_MensajeError = ms_MensajeError + "Por favor elija el Almacén Origen" + Chr(13)
        ElseIf cmbTipoFinanciamiento.Text = "" Then
            ms_MensajeError = ms_MensajeError + "Por favor elija la Fuente de Financiamiento" + Chr(13)
            cmbTipoFinanciamiento.SetFocus
        ElseIf cmbTipoReceta.Text = "" Then
            ms_MensajeError = ms_MensajeError + "Por favor elija el Tipo de Receta" + Chr(13)
            cmbTipoReceta.SetFocus
        End If
   Else
        If cmbAlmOrigen.Text = "" Then
            ms_MensajeError = ms_MensajeError + "Por favor elija el Almacén Origen" + Chr(13)
        ElseIf txtDatosDeCuenta.Text = "" Then
            ms_MensajeError = ms_MensajeError + "Por favor ingrese el N° de Cuenta" + Chr(13)
            txtNcuenta.SetFocus
        ElseIf txtNombrePaciente.Text = "" Then
            ms_MensajeError = ms_MensajeError + "Ese N° de Cuenta no tiene Paciente" + Chr(13)
            txtNhistoria.SetFocus
        ElseIf cmbTipoFinanciamiento.Text = "" Then
            ms_MensajeError = ms_MensajeError + "Por favor elija la Fuente de Financiamiento" + Chr(13)
            cmbTipoFinanciamiento.SetFocus
        ElseIf cmbTipoReceta.Text = "" Then
            ms_MensajeError = ms_MensajeError + "Por favor elija el Tipo de Receta" + Chr(13)
            cmbTipoReceta.SetFocus
        End If
   End If
   lnTotalDocumento = grdProductos.DevuelveTotal
   Set mRs_Productos = grdProductos.DevuelveProductos
   If mRs_Productos.RecordCount = 0 Then
       ms_MensajeError = ms_MensajeError + "Por favor Ingrese Productos" + Chr(13)
   Else
        mRs_Productos.MoveFirst
        Do While Not mRs_Productos.EOF
           If Trim(mRs_Productos.Fields!codigo) = "" Or Trim(mRs_Productos.Fields!nombreProducto) = "" Then
              mRs_Productos.Delete
              mRs_Productos.Update
           ElseIf mRs_Productos.Fields!cantidad <= 0 Or mRs_Productos!cantidad > mRs_Productos!saldo Then
              ms_MensajeError = ms_MensajeError + "El producto " + Trim(mRs_Productos.Fields!codigo) + " - " + Trim(mRs_Productos.Fields!nombreProducto) + "  Tiene problemas de Saldo" + Chr(13)
           ElseIf ml_EsOficinaElTipoFinanciamiento = True And mRs_Productos!precioDelSeguro <= 0 Then
              ms_MensajeError = ms_MensajeError + "El producto " + Trim(mRs_Productos.Fields!codigo) + "  no se encuentra en el Catálogo del Tipo Financiamiento elegido" + Chr(13)
           End If
           mRs_Productos.MoveNext
        Loop
   End If
   If ms_MensajeError <> "" Then
       MsgBox ms_MensajeError, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function


Sub CargaDatosAlObjetosDeDatos()
    Select Case mi_Opcion
    Case sghAgregar
        If optPreventa.Value Then
            With mo_DofarmPreVenta
                .FechaCreacion = txtFregistro.Text
                .HoraCreacion = lcBuscaParametro.RetornaHoraServidorSQL
                .IdAlmacen = Val(mo_cmbAlmacenOrigen.BoundText)
                .idCuentaAtencion = Val(txtNcuenta.Text)
                .idDiagnostico = ml_IdDiagnostico
                .idEstadoPreventa = 1   'Por cancelar en Caja
                .idPaciente = ml_idPaciente
                .idPrescriptor = Val(mo_cmbPrescriptor.BoundText)
                .IdTipoFinanciamiento = Val(mo_cmbTipoFinanciamiento.BoundText)
                .idTipoReceta = Val(mo_cmbTipoReceta.BoundText)
                .idUsuario = ml_idUsuario
                .IdUsuarioAuditoria = ml_idUsuario
                .idVendedor = ml_idUsuario
                .Total = lnTotalDocumento
                If Val(txtRedondeo.Text) > 0 And CCur(txtRedondeo.Text) <> lnTotalDocumento Then
                      .Total = CCur(txtRedondeo.Text)
                End If
            End With
        Else
            With mo_DoFarmMovimiento
                .FechaCreacion = txtFregistro.Text
                .HoraCreacion = lcBuscaParametro.RetornaHoraServidorSQL
                .IdAlmacenDestino = 0   '<<ninguno>>
                .IdAlmacenOrigen = Val(mo_cmbAlmacenOrigen.BoundText)
                .idEstadoMovimiento = 1   'registrado
                .idTipoConcepto = ml_idTipoConcepto
                .idUsuario = ml_idUsuario
                .IdUsuarioAuditoria = ml_idUsuario
                .MovTipo = lcConstanteMovimientoSalida
                .Observaciones = txtObservaciones.Text
                .Total = lnTotalDocumento
            End With
            With mo_DoFarmMovimientoVentas
                .idCuentaAtencion = Val(txtNcuenta.Text)
                .idDiagnostico = ml_IdDiagnostico
                .idPaciente = ml_idPaciente
                .idPrescriptor = Val(mo_cmbPrescriptor.BoundText)
                .IdTipoFinanciamiento = Val(mo_cmbTipoFinanciamiento.BoundText)
                .idTipoReceta = Val(mo_cmbTipoReceta.BoundText)
                .IdUsuarioAuditoria = ml_idUsuario
                .MovTipo = lcConstanteMovimientoSalida
                .tipoVenta = lcConstanteVentaDirecta
                .idFuenteFinanciamiento = ml_IdFuenteFinanciamiento
            End With
        End If
   Case sghModificar
        If optPreventa.Value Then
            With mo_DofarmPreVenta
                .FechaModificacion = lcBuscaParametro.RetornaFechaServidorSQL
                .idDiagnostico = ml_IdDiagnostico
                .idPrescriptor = Val(mo_cmbPrescriptor.BoundText)
                .idTipoReceta = Val(mo_cmbTipoReceta.BoundText)
                .idUsuarioModifica = ml_idUsuario
                .Total = lnTotalDocumento
                .IdUsuarioAuditoria = ml_idUsuario
                If Val(txtRedondeo.Text) > 0 And CCur(txtRedondeo.Text) <> lnTotalDocumento Then
                      .Total = CCur(txtRedondeo.Text)
                End If
            End With
        Else
            With mo_DoFarmMovimiento
                .Observaciones = txtObservaciones.Text
                .Total = lnTotalDocumento
                .IdUsuarioAuditoria = ml_idUsuario
                .FechaCreacion = txtFregistro.Text
            End With
            With mo_DoFarmMovimientoVentas
                .idDiagnostico = ml_IdDiagnostico
                .idPrescriptor = Val(mo_cmbPrescriptor.BoundText)
                .idTipoReceta = Val(mo_cmbTipoReceta.BoundText)
                .IdUsuarioAuditoria = ml_idUsuario
                .idFuenteFinanciamiento = ml_IdFuenteFinanciamiento
            End With
        End If
   Case sghEliminar
        If optPreventa.Value Then
            With mo_DofarmPreVenta
                .idEstadoPreventa = 0   'Anulado
                .IdUsuarioAuditoria = ml_idUsuario
            End With
        Else
            With mo_DoFarmMovimiento
                .fechaAnulacion = lcBuscaParametro.RetornaFechaServidorSQL
                .idEstadoMovimiento = 0   'Anulado
                .IdUsuarioAuditoria = ml_idUsuario
            End With
        End If
   End Select
End Sub

Function AgregarDatos() As Boolean
    If optPreventa.Value Then
        AgregarDatos = mo_ReglasFarmacia.AgregaDatosDePreVenta(mo_DofarmPreVenta, mRs_Productos)
        txtNpreventa.Text = mo_DofarmPreVenta.idPreVenta
    Else
        AgregarDatos = mo_ReglasFarmacia.AgregaDatosDeVentaDirecta(mo_DoFarmMovimiento, mo_DoFarmMovimientoVentas, mRs_Productos, ml_EsOficinaElTipoFinanciamiento)
        txtDocumento.Text = mo_DoFarmMovimiento.DocumentoNumero
    End If
    ms_MensajeError = mo_ReglasFarmacia.MensajeError
End Function
Function ModificarDatos() As Boolean
    If optPreventa.Value Then
        ModificarDatos = mo_ReglasFarmacia.ModificaDatosDePreVenta(mo_DofarmPreVenta, mRs_Productos)
    Else
        ModificarDatos = mo_ReglasFarmacia.ModificaDatosVentaDirecta(mo_DoFarmMovimiento, mo_DoFarmMovimientoVentas, mRs_Productos, ml_EsOficinaElTipoFinanciamiento)
    End If
    ms_MensajeError = mo_ReglasFarmacia.MensajeError
End Function
Function Anular() As Boolean
    If optPreventa.Value Then
        Anular = mo_ReglasFarmacia.AnulaPreVenta(mo_DofarmPreVenta)
    Else
        Anular = mo_ReglasFarmacia.AnulaNotaSalida(mo_DoFarmMovimiento)
    End If
    ms_MensajeError = mo_ReglasFarmacia.MensajeError
End Function
Private Sub btnCancelar_Click()
     Me.Visible = False
     LimpiarVariablesDeMemoria
End Sub




Private Sub txtNhistoria_LostFocus()
      If txtNhistoria.Text <> "" Then
        Dim oConexion As New ADODB.Connection
        Dim oDOPaciente As New SIGHComun.DOPaciente
        oDOPaciente.NroHistoriaClinica = txtNhistoria.Text
        oConexion.Open SIGHComun.CadenaConexion
        Set oBuscaHistoria.Conexion = oConexion
        If oBuscaHistoria.SeleccionarPorHistoriaClinicaDefinitiva(oDOPaciente) Then
           ml_idPaciente = oDOPaciente.idPaciente
           txtNombrePaciente.Text = Trim(oDOPaciente.ApellidoPaterno) & " " & Trim(oDOPaciente.ApellidoMaterno) & " " & oDOPaciente.PrimerNombre
        Else
           ml_idPaciente = 0
           txtNombrePaciente.Text = ""
        End If
      End If

End Sub

Private Sub txtObservaciones_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtObservaciones

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
        Case vbKeyEscape
'           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

Sub LimpiarDatos()
   txtNcuenta.Text = ""
   txtDatosDeCuenta.Text = ""
   txtPlan.Text = ""
   txtNhistoria.Text = ""
   txtNombrePaciente.Text = ""
   ml_idPaciente = 0
   ml_IdFuenteFinanciamiento = 0
   cmbTipoFinanciamiento.Text = ""
   cmbTipoReceta.Text = ""
   txtTurno.Text = ""
   txtCaja.Text = ""
   txtCajero.Text = ""
   ml_IdCajero = 0
   txtVendedor.Text = ""
   ml_IdVendedor = 0
   cmbPrescriptor.Text = ""
   ml_IdDiagnostico = 0
   txtDx.Text = ""
   txtNombreDx.Text = ""
   txtObservaciones.Text = ""
   ml_movNumero = ""
   txtDocumento.Text = ""
   txtNpreventa.Text = ""
   lnTotalDocumento = 0
   grdProductos.movNumero = 0
   grdProductos.LimpiarGrilla
   grdProductos.AgregaRegistro
   If optVentas.Value = True Then
      optVentas_Click 1
      txtNcuenta.SetFocus
   Else
      optPreventa_Click 1
      cmbTipoReceta.SetFocus
   End If
   'Me.KeyPreview = True
End Sub


Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_Formulario = Nothing
    Set mo_cmbAlmacenOrigen = Nothing
    Set mo_cmbPrescriptor = Nothing
    Set mo_cmbTipoFinanciamiento = Nothing
    Set mo_cmbTipoReceta = Nothing
    Set mo_ReglasFarmacia = Nothing
    Set mo_Teclado = Nothing
    Set oRsTipoFinanciamiento = Nothing
    Set mo_AdminServiciosComunes = Nothing
    Set mo_AdminAdmision = Nothing
    Set oBuscaHistoria = Nothing
    Set mo_DofarmPreVenta = Nothing
    Set mo_DoPaciente = Nothing
    Set lcBuscaParametro = Nothing
    Set mRs_Productos = Nothing
    Set mo_DoFarmMovimiento = Nothing
    Set mo_DoFarmMovimientoVentas = Nothing
    Set mo_ReglasCaja = Nothing
    
End Sub


Private Sub txtRedondeo_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
           KeyAscii = 0
       End If
End Sub

Private Sub txtRedondeo_LostFocus()
     If CCur(txtRedondeo.Text) - grdProductos.DevuelveTotal > 0.2 Then
        MsgBox "El redondeo es mayor de 0.20", vbInformation, Me.Caption
        txtRedondeo.Text = grdProductos.DevuelveTotal
        Exit Sub
     End If
     If grdProductos.DevuelveTotal - CCur(txtRedondeo.Text) > 0.2 Then
        MsgBox "El redondeo es menor de 0.20 ", vbInformation, Me.Caption
        txtRedondeo.Text = grdProductos.DevuelveTotal
        Exit Sub
     End If
End Sub
