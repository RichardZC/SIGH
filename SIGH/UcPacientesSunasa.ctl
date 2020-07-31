VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl UcPacientesSunasa 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   5955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11520
   LockControls    =   -1  'True
   ScaleHeight     =   5955
   ScaleWidth      =   11520
   Begin VB.Frame fraDatosPaciente 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5925
      Left            =   30
      TabIndex        =   31
      Top             =   0
      Width           =   11475
      Begin VB.Frame FraChk 
         Height          =   765
         Left            =   10050
         TabIndex        =   80
         Top             =   120
         Width           =   1305
         Begin VB.CheckBox chkNuevoSeguro 
            Alignment       =   1  'Right Justify
            Caption         =   "Nuevo Seguro"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   30
            TabIndex        =   82
            Top             =   150
            Width           =   1215
         End
         Begin VB.CheckBox chkNoTieneSeguro 
            Alignment       =   1  'Right Justify
            Caption         =   "SIN Seguro"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   30
            TabIndex        =   81
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.TextBox txtPais 
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
         Left            =   8310
         MaxLength       =   10
         TabIndex        =   76
         Top             =   510
         Width           =   1725
      End
      Begin VB.TextBox txtDocumento 
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
         Left            =   1500
         MaxLength       =   10
         TabIndex        =   74
         Top             =   510
         Width           =   1845
      End
      Begin VB.TextBox txtNdocumento 
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
         Left            =   5220
         MaxLength       =   10
         TabIndex        =   71
         Top             =   510
         Width           =   1695
      End
      Begin VB.TextBox txtSexo 
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
         Left            =   8310
         TabIndex        =   69
         Top             =   180
         Width           =   1725
      End
      Begin VB.TextBox txtPaciente 
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
         Left            =   1500
         TabIndex        =   67
         Top             =   180
         Width           =   5415
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos del Titular"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Left            =   7020
         TabIndex        =   63
         Top             =   810
         Width           =   4395
         Begin VB.ComboBox cmbPaisTitular 
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
            Left            =   1260
            TabIndex        =   4
            Top             =   270
            Width           =   1770
         End
         Begin VB.TextBox txtNdocumentoTitular 
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
            Left            =   1260
            MaxLength       =   10
            TabIndex        =   6
            Top             =   990
            Width           =   1755
         End
         Begin VB.ComboBox cmdDocumentoTitular 
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
            Left            =   1260
            TabIndex        =   5
            Top             =   630
            Width           =   1785
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "País"
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
            TabIndex        =   66
            Top             =   330
            Width           =   300
         End
         Begin VB.Label Label36 
            Caption         =   "Nº Documento"
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
            TabIndex        =   65
            Top             =   1080
            Width           =   1245
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Doc&umento"
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
            TabIndex        =   64
            Top             =   690
            Width           =   960
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Datos del Paciente (Asegurado)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Left            =   60
         TabIndex        =   58
         Top             =   810
         Width           =   6915
         Begin VB.TextBox txtFnacimiento 
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
            Left            =   1440
            MaxLength       =   10
            TabIndex        =   83
            Top             =   990
            Width           =   1875
         End
         Begin VB.TextBox txtUbigeo 
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
            Left            =   5130
            TabIndex        =   78
            Top             =   1020
            Width           =   1665
         End
         Begin VB.TextBox txtApellidoCasada 
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
            Left            =   1440
            MaxLength       =   35
            TabIndex        =   0
            Top             =   270
            Width           =   1845
         End
         Begin VB.TextBox txtNroDocumentoAnterior 
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
            Left            =   5130
            MaxLength       =   10
            TabIndex        =   3
            Top             =   660
            Width           =   1665
         End
         Begin VB.ComboBox cmbDocumentoAnterior 
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
            Left            =   1440
            TabIndex        =   2
            Top             =   630
            Width           =   1875
         End
         Begin VB.ComboBox cmbParentescoTitular 
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
            Left            =   5130
            TabIndex        =   1
            Top             =   270
            Width           =   1695
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            Caption         =   "Ubigeo (Domicilio)"
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
            Left            =   3210
            TabIndex        =   79
            Top             =   1050
            Width           =   1905
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "F.Nacimiento"
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
            TabIndex        =   77
            Top             =   1050
            Width           =   1050
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Apellido Casada"
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
            TabIndex        =   62
            Top             =   330
            Width           =   1245
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Caption         =   "N° Documento (ant)"
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
            Left            =   3210
            TabIndex        =   61
            Top             =   690
            Width           =   1905
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Docum (ant)"
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
            TabIndex        =   60
            Top             =   690
            Width           =   1050
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Parentes.con Titular"
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
            Left            =   3465
            TabIndex        =   59
            Top             =   330
            Width           =   1650
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Datos del Encargado del Sepelio (SIS)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   60
         TabIndex        =   43
         Top             =   4410
         Width           =   6945
         Begin VB.ComboBox cmbSepelioSexo 
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
            Left            =   1410
            TabIndex        =   23
            Top             =   990
            Width           =   1965
         End
         Begin VB.TextBox txtSepelioApellidosYnombre 
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
            Left            =   1410
            MaxLength       =   100
            TabIndex        =   20
            Top             =   270
            Width           =   5445
         End
         Begin VB.TextBox txtSepelioDNI 
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
            Left            =   5100
            MaxLength       =   8
            TabIndex        =   22
            Top             =   630
            Width           =   1755
         End
         Begin MSMask.MaskEdBox txtSepelioFnacimiento 
            Height          =   315
            Left            =   1410
            TabIndex        =   21
            Top             =   630
            Width           =   1935
            _ExtentX        =   3413
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
         Begin VB.Label lblSepelioEncargado 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Apell. Nombres"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   120
            TabIndex        =   47
            Top             =   300
            Width           =   1230
         End
         Begin VB.Label lblSepelioSexo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   120
            TabIndex        =   46
            Top             =   1050
            Width           =   405
         End
         Begin VB.Label lblSepelioDni 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N° DNI"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   4470
            TabIndex        =   45
            Top             =   660
            Width           =   570
         End
         Begin VB.Label lblSepelioFnacimiento 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "F.Nacimiento"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   120
            TabIndex        =   44
            Top             =   660
            Width           =   1050
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1455
         Left            =   7050
         TabIndex        =   39
         Top             =   4410
         Width           =   4365
         Begin VB.ComboBox cmbTipoOperacion 
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
            Left            =   1380
            TabIndex        =   25
            Top             =   600
            Width           =   1905
         End
         Begin VB.TextBox txtDNIusuario 
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
            Left            =   1380
            MaxLength       =   8
            TabIndex        =   24
            Top             =   240
            Width           =   1785
         End
         Begin MSMask.MaskEdBox txtFechaEnvio 
            Height          =   315
            Left            =   1395
            TabIndex        =   26
            Top             =   990
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtHoraEnvio 
            Height          =   315
            Left            =   2595
            TabIndex        =   27
            Top             =   990
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Envío"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   90
            TabIndex        =   48
            Top             =   1050
            Width           =   975
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "(realiza operación)"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   3150
            TabIndex        =   42
            Top             =   270
            Width           =   1140
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo operación"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   90
            TabIndex        =   41
            Top             =   660
            Width           =   1215
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DNI Usuario "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   90
            TabIndex        =   40
            Top             =   300
            Width           =   1005
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos del Seguro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2145
         Left            =   60
         TabIndex        =   32
         Top             =   2220
         Width           =   11355
         Begin VB.TextBox txtProductoPlan2 
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
            Left            =   2130
            MaxLength       =   4
            TabIndex        =   11
            Top             =   630
            Width           =   1155
         End
         Begin VB.ComboBox cmbEstadoSeguro 
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
            ItemData        =   "UcPacientesSunasa.ctx":0000
            Left            =   1410
            List            =   "UcPacientesSunasa.ctx":000A
            TabIndex        =   18
            Top             =   1710
            Width           =   1905
         End
         Begin VB.ComboBox cmbValidacionRegIden 
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
            ItemData        =   "UcPacientesSunasa.ctx":0020
            Left            =   5160
            List            =   "UcPacientesSunasa.ctx":002A
            TabIndex        =   17
            Top             =   1350
            Width           =   1815
         End
         Begin VB.TextBox txtNroAfiliacion3 
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
            Left            =   5940
            MaxLength       =   8
            TabIndex        =   9
            Top             =   270
            Width           =   1035
         End
         Begin VB.TextBox txtNroAfiliacion2 
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
            Left            =   5550
            MaxLength       =   2
            TabIndex        =   8
            Top             =   270
            Width           =   375
         End
         Begin VB.TextBox txtNroAfiliacion1 
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
            Left            =   5160
            MaxLength       =   10
            TabIndex        =   30
            Top             =   270
            Width           =   375
         End
         Begin VB.TextBox txtCodigoIAFA 
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
            Left            =   5160
            MaxLength       =   5
            TabIndex        =   19
            Top             =   1710
            Width           =   1815
         End
         Begin VB.TextBox txtCarnetIdentidad 
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
            Left            =   8340
            MaxLength       =   10
            TabIndex        =   15
            Top             =   990
            Width           =   1665
         End
         Begin VB.TextBox txtCodEstablecIAFA 
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
            Left            =   8325
            MaxLength       =   8
            TabIndex        =   28
            Top             =   270
            Width           =   1665
         End
         Begin VB.TextBox txtCodEstablecRENAES 
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
            Left            =   8325
            MaxLength       =   8
            TabIndex        =   29
            Top             =   630
            Width           =   1665
         End
         Begin VB.TextBox txtRUCempleador 
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
            Left            =   1410
            MaxLength       =   11
            TabIndex        =   16
            Top             =   1350
            Width           =   1905
         End
         Begin VB.ComboBox cmbRegimen 
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
            Left            =   1410
            TabIndex        =   7
            Top             =   270
            Width           =   1905
         End
         Begin VB.TextBox txtProductoPlan1 
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
            Left            =   1410
            MaxLength       =   3
            TabIndex        =   10
            Top             =   630
            Width           =   705
         End
         Begin VB.ComboBox cmbTipoAfiliacion 
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
            Left            =   5160
            TabIndex        =   12
            Top             =   630
            Width           =   1815
         End
         Begin MSMask.MaskEdBox txtFechaInicioAfiliacion 
            Height          =   315
            Left            =   1410
            TabIndex        =   13
            Top             =   990
            Width           =   1875
            _ExtentX        =   3307
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
         Begin MSMask.MaskEdBox txtFechaFinalAfiliacion 
            Height          =   315
            Left            =   5160
            TabIndex        =   14
            Top             =   990
            Width           =   1815
            _ExtentX        =   3201
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
         Begin VB.Label lblAfiliacionSIS 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Cod.Afiliación (SIS)"
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
            Left            =   3615
            TabIndex        =   57
            Top             =   300
            Width           =   1545
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código IAFA"
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
            Left            =   4170
            TabIndex        =   56
            Top             =   1770
            Width           =   1005
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Validac.Reg.Identidad"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   3390
            TabIndex        =   55
            Top             =   1410
            Width           =   1770
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N° Carnet Ident"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   7020
            TabIndex        =   54
            Top             =   1050
            Width           =   1320
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cód.Establec"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   7290
            TabIndex        =   53
            Top             =   330
            Width           =   1050
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RUC empleador"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   120
            TabIndex        =   52
            Top             =   1380
            Width           =   1245
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "(adscripción IAFA)"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   10035
            TabIndex        =   51
            Top             =   300
            Width           =   1185
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cód.Establec"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   7290
            TabIndex        =   50
            Top             =   660
            Width           =   1050
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "(adscrip.RENAES)"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   10035
            TabIndex        =   49
            Top             =   660
            Width           =   1230
         End
         Begin VB.Label lblRegimen 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Régimen"
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
            TabIndex        =   38
            Top             =   330
            Width           =   705
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "F.Final afiliación"
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
            Left            =   3945
            TabIndex        =   37
            Top             =   1020
            Width           =   1215
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "F.Inicio afiliación"
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
            Top             =   1020
            Width           =   1440
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Producto/Plan"
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
            TabIndex        =   35
            Top             =   660
            Width           =   1155
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Afiliación"
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
            Left            =   3780
            TabIndex        =   34
            Top             =   690
            Width           =   1380
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estado Seguro"
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
            TabIndex        =   33
            Top             =   1740
            Width           =   1200
         End
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "País"
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
         Left            =   7830
         TabIndex        =   75
         Top             =   540
         Width           =   300
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Documento"
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
         TabIndex        =   73
         Top             =   570
         Width           =   960
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
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
         Height          =   225
         Left            =   3960
         TabIndex        =   72
         Top             =   540
         Width           =   1245
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   210
         Left            =   7800
         TabIndex        =   70
         Top             =   210
         Width           =   405
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   240
         TabIndex        =   68
         Top             =   240
         Width           =   705
      End
   End
End
Attribute VB_Name = "UcPacientesSunasa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para registrar datos SUNASA de Paciente
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Teclado As New sighEntidades.Teclado
Dim mo_Formulario As New sighEntidades.Formulario
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim ml_IdPaciente As Long
Dim mo_cmbParentescoTitular As New sighEntidades.ListaDespleglable
Dim mo_cmbDocumentoAnterior As New sighEntidades.ListaDespleglable
Dim mo_cmbTipoOperacion As New sighEntidades.ListaDespleglable
Dim mo_cmbTipoAfiliacion As New sighEntidades.ListaDespleglable
Dim mo_cmbRegimen As New sighEntidades.ListaDespleglable
Dim mo_cmbPaisTitular As New sighEntidades.ListaDespleglable
Dim mo_cmdDocumentoTitular As New sighEntidades.ListaDespleglable
Dim mo_cmbSepelioSexo As New sighEntidades.ListaDespleglable
Dim mo_AdminServiciosGeograficos As New SIGHNegocios.ReglasServGeograf
Dim mo_AdminServiciosComunes As New ReglasComunes
Dim mo_AdminAdmision As New ReglasAdmision
Dim mi_idTipoFinanciamiento As sghComoSeTrabajaEnEstadoCuentaLosSeguros
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_idSunasaPacienteHistorico As Long
Dim ml_idUsuario As Long
Dim mb_YaNoTieneSeguroUltimoRegistroGrabado As Boolean
Public Event SePresionoTeclaEspecial(KeyCode As Integer)


Property Let idTipoFinanciamiento(iValue As sghComoSeTrabajaEnEstadoCuentaLosSeguros)
   mi_idTipoFinanciamiento = iValue
   CambiaDeTipoDeFinanciamiento
End Property

Property Let Opcion(iValue As sghOpciones)
   mi_Opcion = iValue
End Property
Property Let idPaciente(lValue As Long)
   ml_IdPaciente = lValue
End Property
Property Get idSunasaPacienteHistorico() As Long
   idSunasaPacienteHistorico = ml_idSunasaPacienteHistorico
End Property

Property Let idSunasaPacienteHistorico(lValue As Long)
   ml_idSunasaPacienteHistorico = lValue
End Property

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property


Sub Inicializar()
    'carga combos
    Set mo_cmbParentescoTitular.MiComboBox = cmbParentescoTitular
    mo_cmbParentescoTitular.BoundColumn = "IdParentesco"
    mo_cmbParentescoTitular.ListField = "Parentesco"
    Set mo_cmbParentescoTitular.RowSource = mo_AdminServiciosComunes.SunasaTiposParentescoSeleccionarTodos
    '
    Set mo_cmbTipoOperacion.MiComboBox = cmbTipoOperacion
    mo_cmbTipoOperacion.BoundColumn = "IdOperacion"
    mo_cmbTipoOperacion.ListField = "Operacion"
    Set mo_cmbTipoOperacion.RowSource = mo_AdminServiciosComunes.SunasaTiposOperacionSeleccionarTodos
    '
    Set mo_cmbTipoAfiliacion.MiComboBox = cmbTipoAfiliacion
    mo_cmbTipoAfiliacion.BoundColumn = "IdAfiliacion"
    mo_cmbTipoAfiliacion.ListField = "Afiliacion"
    Set mo_cmbTipoAfiliacion.RowSource = mo_AdminServiciosComunes.SunasaTiposAfiliacionSeleccionarTodos
    '
    Set mo_cmbRegimen.MiComboBox = cmbRegimen
    mo_cmbRegimen.BoundColumn = "IdRegimen"
    mo_cmbRegimen.ListField = "Regimen"
    Set mo_cmbRegimen.RowSource = mo_AdminServiciosComunes.SunasaTiposRegimenSeleccionarTodos
    '
    Set mo_cmbDocumentoAnterior.MiComboBox = cmbDocumentoAnterior
    mo_cmbDocumentoAnterior.BoundColumn = "IdDocIdentidad"
    mo_cmbDocumentoAnterior.ListField = "DescripcionLarga"
    Set mo_cmbDocumentoAnterior.RowSource = mo_AdminServiciosComunes.TiposDocIdentidadSeleccionarTodos()
    '
    Set mo_cmdDocumentoTitular.MiComboBox = cmdDocumentoTitular
    mo_cmdDocumentoTitular.BoundColumn = "IdDocIdentidad"
    mo_cmdDocumentoTitular.ListField = "DescripcionLarga"
    Set mo_cmdDocumentoTitular.RowSource = mo_AdminServiciosComunes.TiposDocIdentidadSeleccionarTodos()
    '
    Set mo_cmbPaisTitular.MiComboBox = cmbPaisTitular
    mo_cmbPaisTitular.BoundColumn = "IdPais"
    mo_cmbPaisTitular.ListField = "Nombre"
    Set mo_cmbPaisTitular.RowSource = mo_AdminServiciosGeograficos.PaisesSeleccionarTodos()
    PaisTitularDefault
    '
    Set mo_cmbSepelioSexo.MiComboBox = cmbSepelioSexo
    mo_cmbSepelioSexo.BoundColumn = "IdtipoSexo"
    mo_cmbSepelioSexo.ListField = "DescripcionLarga"
    Set mo_cmbSepelioSexo.RowSource = mo_AdminServiciosComunes.TiposSexoSeleccionarTodos()
    '
    mo_Formulario.HabilitarDeshabilitar txtNroAfiliacion1, False
    mo_Formulario.HabilitarDeshabilitar txtPaciente, False
    mo_Formulario.HabilitarDeshabilitar txtSexo, False
    mo_Formulario.HabilitarDeshabilitar txtDocumento, False
    mo_Formulario.HabilitarDeshabilitar txtNdocumento, False
    mo_Formulario.HabilitarDeshabilitar txtPais, False
    mo_Formulario.HabilitarDeshabilitar txtFnacimiento, False
    mo_Formulario.HabilitarDeshabilitar txtUbigeo, False
    mo_Formulario.HabilitarDeshabilitar txtCodEstablecIAFA, False
    mo_Formulario.HabilitarDeshabilitar txtCodEstablecRENAES, False
    '
    txtNroAfiliacion1.Text = lcBuscaParametro.SeleccionaFilaParametro(239) 'codigo de la DISA
    '
    CambiaDeTipoDeFinanciamiento
End Sub

Sub CambiaDeTipoDeFinanciamiento()
    lblSepelioEncargado.ForeColor = vbBlack
    lblSepelioDni.ForeColor = vbBlack
    lblSepelioFnacimiento.ForeColor = vbBlack
    lblSepelioSexo.ForeColor = vbBlack
    lblRegimen.ForeColor = vbBlack
    lblAfiliacionSIS.ForeColor = vbBlack
    Select Case mi_idTipoFinanciamiento
    Case sghTrabajaSeguroSIS
        lblSepelioEncargado.ForeColor = vbBlue
        lblSepelioDni.ForeColor = vbBlue
        lblSepelioFnacimiento.ForeColor = vbBlue
        lblSepelioSexo.ForeColor = vbBlue
        lblRegimen.ForeColor = vbBlue
        lblAfiliacionSIS.ForeColor = vbBlue
    End Select

End Sub

Public Sub PaisTitularDefault()
    If mo_cmbPaisTitular.BoundText = "" Then
        mo_cmbPaisTitular.BoundText = "166" 'Peru
    End If
End Sub

Private Sub chkNoTieneSeguro_Click()
    If chkNoTieneSeguro.Value = 1 Then
       chkNuevoSeguro.Value = 0
       LimpiarDatos
       InhabilitaHabilitaControles False
    Else
       InhabilitaHabilitaControles True
    End If
End Sub

Private Sub chkNuevoSeguro_Click()
   If chkNuevoSeguro.Value = 1 Then
      chkNoTieneSeguro.Value = 0
      LimpiarDatos
      InhabilitaHabilitaControles True
      txtCodEstablecIAFA.Text = lcBuscaParametro.SeleccionaFilaParametro(280) 'codigo de HOspital segun RENAES
      txtCodEstablecRENAES.Text = lcBuscaParametro.SeleccionaFilaParametro(280) 'codigo de HOspital segun RENAES
      '
      SetFocusEnApellidoCasada
   End If
End Sub

Private Sub cmbDocumentoAnterior_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbDocumentoAnterior
   RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub



Private Sub cmbEstadoSeguro_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbEstadoSeguro
   RaiseEvent SePresionoTeclaEspecial(KeyCode)

End Sub

Private Sub cmbPaisTitular_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbPaisTitular
   RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub cmbParentescoTitular_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbParentescoTitular
   RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub




Private Sub cmbRegimen_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbRegimen
   RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub


Private Sub cmbSepelioSexo_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbSepelioSexo
   RaiseEvent SePresionoTeclaEspecial(KeyCode)

End Sub

Private Sub cmbTipoAfiliacion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbTipoAfiliacion
   RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub cmbTipoOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbTipoOperacion
   RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub



Private Sub cmbValidacionRegIden_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbValidacionRegIden
   RaiseEvent SePresionoTeclaEspecial(KeyCode)

End Sub

Private Sub cmdDocumentoTitular_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmdDocumentoTitular
   RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub








Private Sub txtApellidoCasada_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtApellidoCasada
   RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub





Private Sub txtCarnetIdentidad_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtCarnetIdentidad
   RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub


Private Sub txtCodEstablecIAFA_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtCodEstablecIAFA
   RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub


Private Sub txtCodEstablecRENAES_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtCodEstablecRENAES
   RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtCodigoIAFA_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtCodigoIAFA
   RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtDNIusuario_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtDNIusuario
   RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub






Private Sub txtFechaEnvio_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFechaEnvio
   RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtFechaEnvio_LostFocus()
If Not EsFecha(txtFechaEnvio.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFechaEnvio.Text = sighEntidades.FECHA_VACIA_DMY
        Exit Sub
    End If
End Sub

Private Sub txtFechaFinalAfiliacion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFechaFinalAfiliacion
   RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtFechaFinalAfiliacion_LostFocus()
If Not EsFecha(txtFechaFinalAfiliacion.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFechaFinalAfiliacion.Text = sighEntidades.FECHA_VACIA_DMY
        Exit Sub
    End If
End Sub

Private Sub txtFechaInicioAfiliacion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFechaInicioAfiliacion
   RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub



Private Sub txtFechaInicioAfiliacion_LostFocus()
If Not EsFecha(txtFechaInicioAfiliacion.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFechaInicioAfiliacion.Text = sighEntidades.FECHA_VACIA_DMY
        Exit Sub
    End If
End Sub

Private Sub txtHoraEnvio_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtHoraEnvio
   RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtHoraEnvio_LostFocus()
If Not sighEntidades.ValidaHora(txtHoraEnvio.Text) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, ""
            txtHoraEnvio.Text = sighEntidades.HORA_VACIA_HM
        End If
End Sub

Private Sub txtNdocumentoTitular_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNdocumentoTitular
   RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub


Private Sub txtNroAfiliacion2_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNroAfiliacion2
   RaiseEvent SePresionoTeclaEspecial(KeyCode)

End Sub



Private Sub txtNroAfiliacion3_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNroAfiliacion3
   RaiseEvent SePresionoTeclaEspecial(KeyCode)

End Sub

Private Sub txtNroDocumentoAnterior_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNroDocumentoAnterior
   RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub






Private Sub txtProductoPlan1_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtProductoPlan1
   RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub


Private Sub txtProductoPlan2_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtProductoPlan2
   RaiseEvent SePresionoTeclaEspecial(KeyCode)

End Sub

Private Sub txtRUCempleador_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtRUCempleador
   RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub



Private Sub txtSepelioApellidosYnombre_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtSepelioApellidosYnombre
   RaiseEvent SePresionoTeclaEspecial(KeyCode)

End Sub


Private Sub txtSepelioDNI_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtSepelioDNI
   RaiseEvent SePresionoTeclaEspecial(KeyCode)

End Sub


Private Sub txtSepelioFnacimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtSepelioFnacimiento
   RaiseEvent SePresionoTeclaEspecial(KeyCode)

End Sub





Public Sub LimpiarDatos()
    cmbRegimen.Text = ""
    txtNroAfiliacion2.Text = ""
    txtNroAfiliacion3.Text = ""
    txtCodEstablecIAFA.Text = ""
    txtProductoPlan1.Text = ""
    txtProductoPlan2.Text = ""
    cmbTipoAfiliacion.Text = ""
    txtCodEstablecRENAES.Text = ""
    txtFechaInicioAfiliacion.Text = sighEntidades.FECHA_VACIA_DMY
    txtFechaFinalAfiliacion.Text = sighEntidades.FECHA_VACIA_DMY
    txtCodigoIAFA.Text = ""
    txtRUCempleador.Text = ""
    cmbValidacionRegIden.Text = ""
    txtCarnetIdentidad.Text = ""
    txtApellidoCasada.Text = ""
    cmbParentescoTitular.Text = ""
    cmbPaisTitular.Text = ""
    cmbDocumentoAnterior.Text = ""
    txtNroDocumentoAnterior.Text = ""
    cmbPaisTitular.Text = ""
    cmdDocumentoTitular.Text = ""
    txtNdocumentoTitular.Text = ""
    txtSepelioApellidosYnombre.Text = ""
    txtSepelioFnacimiento.Text = sighEntidades.FECHA_VACIA_DMY
    txtSepelioDNI.Text = ""
    cmbSepelioSexo.Text = ""
    txtDNIusuario.Text = ""
    cmbTipoOperacion.Text = ""
    txtFechaEnvio.Text = sighEntidades.FECHA_VACIA_DMY
    txtHoraEnvio.Text = sighEntidades.HORA_VACIA_HMS
'    txtPaciente.Text = ""
'    txtSexo.Text = ""
'    txtDocumento.Text = ""
'    txtNdocumento.Text = ""
'    txtPais.Text = ""
'    txtFnacimiento.Text = ""
'    txtUbigeo.Text = ""
End Sub

Public Sub SetFocusEnRegimen()
   On Error Resume Next
   cmbRegimen.SetFocus
End Sub

Public Function CargarDatosAlObjetoDatos(oDoSunasaPacientesHistoricos As DoSunasaPacientesHistoricos)
    
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE SUNASA
    '---------------------------------------------------------------------------------
   With oDoSunasaPacientesHistoricos
        .AnteriorIdTipoDocumentoAsegurado = Val(mo_cmbDocumentoAnterior.BoundText)
        .AnteriorNroDocumentoAsegurado = txtNroDocumentoAnterior.Text
        .ApellidoCasada = txtApellidoCasada.Text
        .CodigoEstablecimientoIAFA = txtCodEstablecIAFA.Text
        .CodigoEstablecimientoRENAES = txtCodEstablecRENAES.Text
        .CodigoIAFA = txtCodigoIAFA.Text
        .DNIusarioOperacion = txtDNIusuario.Text
        .EstadoDelSeguro = IIf(cmbEstadoSeguro.Text = "", 0, cmbEstadoSeguro.ListIndex + 1)
        If txtFechaEnvio.Text <> sighEntidades.FECHA_VACIA_DMY And txtHoraEnvio.Text <> sighEntidades.HORA_VACIA_HM Then
           .FechaEnvio = CDate(txtFechaEnvio.Text & " " & txtHoraEnvio.Text)
        Else
           .FechaEnvio = 0
        End If
        If txtFechaFinalAfiliacion.Text <> sighEntidades.FECHA_VACIA_DMY Then
           .FechaFinalAfiliacion = CDate(txtFechaFinalAfiliacion.Text)
        Else
           .FechaFinalAfiliacion = 0
        End If
        If txtFechaInicioAfiliacion.Text <> sighEntidades.FECHA_VACIA_DMY Then
           .FechaInicioAfiliacion = CDate(txtFechaInicioAfiliacion.Text)
        Else
           .FechaInicioAfiliacion = 0
        End If
        .IdAfiliacion = Val(mo_cmbTipoAfiliacion.BoundText)
        .idOperacion = Val(mo_cmbTipoOperacion.BoundText)
        .idPaciente = ml_IdPaciente
        .idPaisTitular = Val(mo_cmbPaisTitular.BoundText)
        .idParentesco = Val(mo_cmbParentescoTitular.BoundText)
        .idRegimen = Val(mo_cmbRegimen.BoundText)
        .idSunasaPacienteHistorico = ml_idSunasaPacienteHistorico
        .idTipoDocumentoTitular = Val(mo_cmdDocumentoTitular.BoundText)
        .IdUsuarioAuditoria = ml_idUsuario
        .NroCarnetIdentidad = txtCarnetIdentidad.Text
        .NroDocumentoTitular = txtNdocumentoTitular.Text
        If txtProductoPlan1.Text <> "" And txtProductoPlan2.Text <> "" Then
           .ProductoYplan = Right("     " & txtProductoPlan1.Text, 3) & txtProductoPlan2.Text
        Else
           .ProductoYplan = ""
        End If
        .RUCempleador = txtRUCempleador.Text
        If txtNroAfiliacion1.Text <> "" And txtNroAfiliacion2.Text <> "" And txtNroAfiliacion3.Text <> "" Then
           .SisNroAfiliacion = txtNroAfiliacion1.Text & "-" & txtNroAfiliacion2.Text & "-" & txtNroAfiliacion3.Text
        Else
           .SisNroAfiliacion = ""
        End If
        .SisSepelioDni = txtSepelioDNI.Text
        If txtSepelioFnacimiento.Text <> sighEntidades.FECHA_VACIA_DMY Then
           .SisSepelioFnacimiento = CDate(txtSepelioFnacimiento.Text)
        Else
           .SisSepelioFnacimiento = 0
        End If
        .SisSepelioParienteEncargado = txtSepelioApellidosYnombre.Text
        If cmbSepelioSexo.Text <> "" Then
           .SisSepelioSexo = mo_cmbSepelioSexo.BoundText + 1
        Else
           .SisSepelioSexo = 0
        End If
        If cmbValidacionRegIden.Text = "" Then
           .ValidacionRegIdentidad = 0
        Else
           .ValidacionRegIdentidad = cmbValidacionRegIden.ListIndex
        End If
        .YaNoTieneSeguro = chkNoTieneSeguro.Value
        .NuevoSeguro = chkNuevoSeguro.Value      'No se graba en la BD
        '
        If mi_Opcion = sghModificar Then
           If mb_YaNoTieneSeguroUltimoRegistroGrabado = False And chkNoTieneSeguro.Value = 1 Then
              .NuevoSeguro = True
           End If
        End If
        '
   End With
End Function

Sub CargarDatos(oDoSunasaPacientesHistoricos As DoSunasaPacientesHistoricos)
        If oDoSunasaPacientesHistoricos.idPaciente = 0 Then
            chkNoTieneSeguro.Value = 1
            chkNoTieneSeguro_Click
        Else
            Dim lnPos1 As Integer, lnPos2 As Integer
            With oDoSunasaPacientesHistoricos
                mo_cmbDocumentoAnterior.BoundText = .AnteriorIdTipoDocumentoAsegurado
                txtNroDocumentoAnterior.Text = .AnteriorNroDocumentoAsegurado
                txtApellidoCasada.Text = .ApellidoCasada
                txtCodEstablecIAFA.Text = .CodigoEstablecimientoIAFA
                txtCodEstablecRENAES.Text = .CodigoEstablecimientoRENAES
                txtCodigoIAFA.Text = .CodigoIAFA
                txtDNIusuario.Text = .DNIusarioOperacion
                If .EstadoDelSeguro > 0 Then
                   cmbEstadoSeguro.ListIndex = .EstadoDelSeguro - 1
                End If
                If sighEntidades.EsFecha(CDate(Format(.FechaEnvio, "dd/mm/yyyy")), "DD/MM/AAAA") Then
                   txtFechaEnvio.Text = Format(.FechaEnvio, sighEntidades.DevuelveFechaSoloFormato_DMY)
                   UserControl.txtHoraEnvio.Text = Format(.FechaEnvio, sighEntidades.DevuelveHoraSoloFormato_HMS)
                End If
                If sighEntidades.EsFecha(.FechaFinalAfiliacion, "DD/MM/AAAA") Then
                   txtFechaFinalAfiliacion.Text = Format(.FechaFinalAfiliacion, sighEntidades.DevuelveFechaSoloFormato_DMY)
                End If
                If sighEntidades.EsFecha(.FechaInicioAfiliacion, "DD/MM/AAAA") Then
                   txtFechaInicioAfiliacion.Text = Format(.FechaInicioAfiliacion, sighEntidades.DevuelveFechaSoloFormato_DMY)
                End If
                mo_cmbTipoAfiliacion.BoundText = .IdAfiliacion
                mo_cmbTipoOperacion.BoundText = .idOperacion
                '.idPaciente
                mo_cmbPaisTitular.BoundText = .idPaisTitular
                mo_cmbParentescoTitular.BoundText = .idParentesco
                mo_cmbRegimen.BoundText = .idRegimen
                ml_idSunasaPacienteHistorico = .idSunasaPacienteHistorico
                mo_cmdDocumentoTitular.BoundText = .idTipoDocumentoTitular
                '.IdUsuarioAuditoria
                txtCarnetIdentidad.Text = .NroCarnetIdentidad
                txtNdocumentoTitular.Text = .NroDocumentoTitular
                If .ProductoYplan <> "" Then
                   txtProductoPlan1.Text = Left(.ProductoYplan, 3)
                   txtProductoPlan2.Text = Mid(.ProductoYplan, 4, 4)
                End If
                txtRUCempleador.Text = .RUCempleador
                If .SisNroAfiliacion <> "" Then
                   lnPos1 = InStr(.SisNroAfiliacion, "-")
                   lnPos2 = InStr(lnPos1 + 1, .SisNroAfiliacion, "-")
                   txtNroAfiliacion1.Text = Left(.SisNroAfiliacion, lnPos1 - 1)
                   txtNroAfiliacion2.Text = Mid(.SisNroAfiliacion, lnPos1 + 1, lnPos2 - lnPos1 - 1)
                   txtNroAfiliacion3.Text = Mid(.SisNroAfiliacion, lnPos2 + 1, 100)
                End If
                txtSepelioDNI.Text = .SisSepelioDni
                If sighEntidades.EsFecha(.SisSepelioFnacimiento, "DD/MM/AAAA") Then
                   txtSepelioFnacimiento.Text = Format(.SisSepelioFnacimiento, sighEntidades.DevuelveFechaSoloFormato_DMY)
                End If
                txtSepelioApellidosYnombre.Text = .SisSepelioParienteEncargado
                If Not IsNull(.SisSepelioSexo) Then
                   mo_cmbSepelioSexo.BoundText = .SisSepelioSexo - 1
                End If
                cmbValidacionRegIden.ListIndex = IIf(.ValidacionRegIdentidad = True, 1, 0)
                '
                mb_YaNoTieneSeguroUltimoRegistroGrabado = .YaNoTieneSeguro
                chkNoTieneSeguro.Value = IIf(.YaNoTieneSeguro = True, 1, 0)
                chkNoTieneSeguro_Click
                '
            End With
        End If
End Sub

Public Sub CargarDatosDelUltimoSeguroDelPacienteALosControles(oConexion As Connection)
On Error GoTo ErrrCargaDatos
Dim oDoSunasaPacientesHistoricos As New DoSunasaPacientesHistoricos
        'CARGAR DATOS DE SUNASA
        Set oDoSunasaPacientesHistoricos = mo_AdminAdmision.SunasaPacientesHistoricosSeleccionarPorIdPaciente(ml_IdPaciente, oConexion)
        If mo_AdminAdmision.MensajeError <> "" Then
             MsgBox "No se pudo obtener los datos" + Chr(13) + mo_AdminAdmision.MensajeError, vbInformation, "Datos de SUNASA"
             Exit Sub
        End If
        CargarDatos oDoSunasaPacientesHistoricos
ErrrCargaDatos:
End Sub




Public Sub InhabilitaHabilitaControles(lbValor As Boolean)
    mo_Formulario.HabilitarDeshabilitar cmbRegimen, lbValor
    mo_Formulario.HabilitarDeshabilitar txtNroAfiliacion2, lbValor
    mo_Formulario.HabilitarDeshabilitar txtNroAfiliacion3, lbValor
    'mo_Formulario.HabilitarDeshabilitar txtCodEstablecIAFA, lbValor
    'mo_Formulario.HabilitarDeshabilitar txtCodEstablecRENAES, lbValor
    mo_Formulario.HabilitarDeshabilitar txtProductoPlan1, lbValor
    mo_Formulario.HabilitarDeshabilitar txtProductoPlan2, lbValor
    mo_Formulario.HabilitarDeshabilitar cmbTipoAfiliacion, lbValor
    
    mo_Formulario.HabilitarDeshabilitar txtFechaInicioAfiliacion, lbValor
    mo_Formulario.HabilitarDeshabilitar txtFechaFinalAfiliacion, lbValor
    mo_Formulario.HabilitarDeshabilitar txtCodigoIAFA, lbValor
    mo_Formulario.HabilitarDeshabilitar txtRUCempleador, lbValor
    mo_Formulario.HabilitarDeshabilitar cmbValidacionRegIden, lbValor
    mo_Formulario.HabilitarDeshabilitar txtCarnetIdentidad, lbValor
    mo_Formulario.HabilitarDeshabilitar txtApellidoCasada, lbValor
    mo_Formulario.HabilitarDeshabilitar cmbParentescoTitular, lbValor
    mo_Formulario.HabilitarDeshabilitar cmbPaisTitular, lbValor
    mo_Formulario.HabilitarDeshabilitar cmbDocumentoAnterior, lbValor
    mo_Formulario.HabilitarDeshabilitar txtNdocumentoTitular, lbValor
    mo_Formulario.HabilitarDeshabilitar txtSepelioApellidosYnombre, lbValor
    mo_Formulario.HabilitarDeshabilitar txtSepelioFnacimiento, lbValor
    mo_Formulario.HabilitarDeshabilitar txtSepelioDNI, lbValor
    mo_Formulario.HabilitarDeshabilitar cmbSepelioSexo, lbValor
    mo_Formulario.HabilitarDeshabilitar txtDNIusuario, lbValor
    mo_Formulario.HabilitarDeshabilitar cmbTipoOperacion, lbValor
    mo_Formulario.HabilitarDeshabilitar txtFechaEnvio, lbValor
    mo_Formulario.HabilitarDeshabilitar txtHoraEnvio, lbValor
    mo_Formulario.HabilitarDeshabilitar cmbEstadoSeguro, lbValor
    mo_Formulario.HabilitarDeshabilitar txtNroDocumentoAnterior, lbValor
    mo_Formulario.HabilitarDeshabilitar cmdDocumentoTitular, lbValor
End Sub

Public Sub NuevoSeguro()
    Select Case mi_Opcion
    Case sghModificar
    Case sghAgregar
        chkNuevoSeguro.Value = 1
        chkNuevoSeguro_Click
    Case Else
        FraChk.Enabled = False
    End Select
End Sub

Public Sub YaNoTieneSeguro()
    Select Case mi_Opcion
    Case sghAgregar, sghModificar
        chkNoTieneSeguro.Value = 1
        chkNoTieneSeguro_Click
    Case Else
        FraChk.Enabled = False
    End Select
End Sub

Public Sub HabilitaFrame(lbValor As Boolean)
    FraChk.Enabled = lbValor
End Sub


Public Sub CargarDatosPorId()
On Error GoTo ErrrCargaDatos
Dim oDoSunasaPacientesHistoricos As New DoSunasaPacientesHistoricos
        'CARGAR DATOS DE SUNASA
        Set oDoSunasaPacientesHistoricos = mo_AdminAdmision.SunasaPacientesHistoricosSeleccionarPorId(ml_idSunasaPacienteHistorico)
        If mo_AdminAdmision.MensajeError <> "" Then
             MsgBox "No se pudo obtener los datos" + Chr(13) + mo_AdminAdmision.MensajeError, vbInformation, "Datos de SUNASA"
             Exit Sub
        End If
        CargarDatos oDoSunasaPacientesHistoricos
ErrrCargaDatos:
End Sub

Public Sub idSunasaPacienteHistorico_idPaciente_ConValorCero()
   ml_idSunasaPacienteHistorico = 0
   ml_IdPaciente = 0
End Sub

Public Sub DatosDeCabecera(lcPaciente As String, lcSexo As String, lcDocumento As String, lcNdocumento As String, lcPais As String, lcFnacimiento As String, lcUbigeo As String)
    txtPaciente.Text = lcPaciente
    txtSexo.Text = lcSexo
    txtDocumento.Text = lcDocumento
    txtNdocumento.Text = lcNdocumento
    txtPais.Text = lcPais
    txtFnacimiento.Text = lcFnacimiento
    txtUbigeo.Text = lcUbigeo
End Sub
Public Sub SetFocusEnApellidoCasada()
   On Error Resume Next
   txtApellidoCasada.SetFocus
End Sub

Private Sub txtSepelioFnacimiento_LostFocus()
If Not EsFecha(txtSepelioFnacimiento.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtSepelioFnacimiento.Text = sighEntidades.FECHA_VACIA_DMY
        Exit Sub
    End If
End Sub
